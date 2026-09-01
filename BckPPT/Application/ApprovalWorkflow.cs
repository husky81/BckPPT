using System.IO;
using System.Text;
using BckPPT.Agent;
using BckPPT.Services;
using BckPPT.PowerPoint;
using BckPPT.Rag;

namespace BckPPT.Application;

public sealed class ApprovalWorkflow(
    IPptComService pptService, RenderValidator renderValidator, LayoutDecisionRouter layoutRouter)
{
    public PptJob Create(string instruction, PptTaskPlan plan, DocumentSession session)
    {
        var path = session.Presentation?.Path
            ?? throw new InvalidOperationException("먼저 대상 PPT를 선택해 주세요.");
        if (!File.Exists(path)) throw new FileNotFoundException("대상 PPT 파일을 찾을 수 없습니다.", path);
        return new PptJob
        {
            Id = Guid.NewGuid(), SourcePath = Path.GetFullPath(path),
            SourceLastWriteUtc = File.GetLastWriteTimeUtc(path), Instruction = instruction, Plan = plan
        };
    }

    public async Task<string> PreparePreviewAsync(PptJob job, CancellationToken cancellationToken = default)
    {
        EnsureState(job, PptJobState.Planned);
        EnsureSourceUnchanged(job);
        var lines = new List<string> { $"원본은 변경하지 않습니다: {Path.GetFileName(job.SourcePath)}", "적용 시 사본을 생성합니다." };
        foreach (var command in job.Plan.Commands)
        {
            ValidateCommand(command);
            lines.Add(Describe(command));
            if (command.Name == "replace_shape_text" && command.ExpectedText is not null)
            {
                var shape = pptService.GetShapeInfo(job.SourcePath, command.SlideIndex!.Value, command.ShapeName!);
                if (shape is null || !string.Equals(shape.FullText, command.ExpectedText, StringComparison.Ordinal))
                    throw new InvalidOperationException($"{command.SlideIndex}쪽 '{command.ShapeName}'의 현재 텍스트가 계획의 사전상태와 다릅니다.");
            }
        }
        if (job.Plan.Commands.Any(command => command.Name == "search_template_library"))
        {
            var decision = await layoutRouter.DecideAsync(job.Instruction, job.SourcePath, cancellationToken);
            job.LayoutDecision = decision;
            lines.Add(DescribeDecision(decision));
        }
        job.BeforeThumbnails = await renderValidator.RenderSlidesAsync(
            job.SourcePath, GetSourceSlides(job.Plan), cancellationToken);
        job.Preview = string.Join(Environment.NewLine, lines);
        job.State = PptJobState.PreviewReady;
        return job.Preview;
    }

    public async Task<string> ApplyAsync(PptJob job, CancellationToken cancellationToken = default)
    {
        EnsureState(job, PptJobState.PreviewReady);
        EnsureSourceUnchanged(job);
        job.State = PptJobState.Applying;
        var resultPath = CreateCopyPath(job.SourcePath);
        try
        {
            File.Copy(job.SourcePath, resultPath, overwrite: false);
            foreach (var command in job.Plan.Commands)
            {
                cancellationToken.ThrowIfCancellationRequested();
                switch (command.Name)
                {
                    case "duplicate_slide":
                        await Task.Run(() => pptService.DuplicateSlide(
                            resultPath, command.SourceSlideIndex!.Value, command.InsertAfterSlideIndex!.Value), cancellationToken);
                        break;
                    case "insert_slide" when job.LayoutDecision?.Candidate is { } candidate:
                        if (job.LayoutDecision.Kind == LayoutDecisionKind.TemplateLibrary)
                            await Task.Run(() => pptService.ImportSlide(candidate.FilePath, candidate.SlideIndex,
                                resultPath, command.InsertAfterSlideIndex!.Value), cancellationToken);
                        else
                            await Task.Run(() => pptService.DuplicateSlide(resultPath, candidate.SlideIndex,
                                command.InsertAfterSlideIndex!.Value), cancellationToken);
                        break;
                    case "insert_slide" when command.SourceSlideIndex is not null:
                        await Task.Run(() => pptService.DuplicateSlide(resultPath, command.SourceSlideIndex.Value,
                            command.InsertAfterSlideIndex!.Value), cancellationToken);
                        break;
                    case "replace_shape_text":
                        await Task.Run(() => pptService.SetShapeText(
                            resultPath, command.SlideIndex!.Value, command.ShapeName!, command.NewText!), cancellationToken);
                        break;
                    case "validate_layout":
                        await Task.Run(() => pptService.AnalyzePresentation(resultPath, cancellationToken: cancellationToken), cancellationToken);
                        break;
                }
            }
            var targetSlides = GetResultSlides(job.Plan);
            job.Validation = await renderValidator.ValidateAndCorrectAsync(resultPath, targetSlides, cancellationToken);
            job.AfterThumbnails = await renderValidator.RenderSlidesAsync(resultPath, targetSlides, cancellationToken);
            job.ResultPath = resultPath;
            job.State = PptJobState.Completed;
            return resultPath;
        }
        catch (OperationCanceledException)
        {
            job.State = PptJobState.Cancelled;
            TryDeleteIncompleteCopy(resultPath);
            throw;
        }
        catch (Exception ex)
        {
            job.State = PptJobState.Failed;
            job.Error = ex.Message;
            TryDeleteIncompleteCopy(resultPath);
            throw;
        }
    }

    public void Cancel(PptJob job)
    {
        if (job.State is PptJobState.Applying or PptJobState.Completed) return;
        job.State = PptJobState.Cancelled;
    }

    private static void ValidateCommand(PptPlanCommand command)
    {
        if (command.Name is "duplicate_slide" or "insert_slide" &&
            (command.SourceSlideIndex is null or <= 0 || command.InsertAfterSlideIndex is null or < 0))
            throw new InvalidOperationException($"{command.Name}에 복제 원본/삽입 위치가 없습니다.");
        if (command.Name == "replace_shape_text" &&
            (command.SlideIndex is null or <= 0 || string.IsNullOrWhiteSpace(command.ShapeName) || command.NewText is null))
            throw new InvalidOperationException("replace_shape_text에 슬라이드, Shape 또는 새 텍스트가 없습니다.");
    }

    private static string Describe(PptPlanCommand command) => command.Name switch
    {
        "duplicate_slide" or "insert_slide" => $"• {command.SourceSlideIndex}쪽을 복제해 {command.InsertAfterSlideIndex}쪽 뒤에 삽입",
        "replace_shape_text" => $"• {command.SlideIndex}쪽 '{command.ShapeName}' 텍스트 교체",
        "validate_layout" => "• 적용 결과 레이아웃 검사",
        "save_copy" => "• 결과를 사본으로 저장",
        "analyze_presentation" => $"• {(command.SlideIndex is null ? "문서" : $"{command.SlideIndex}쪽")} 구조 분석",
        "search_template_library" => "• 템플릿 후보 검색(현재 단계에서는 기록만 수행)",
        _ => $"• {command.Name}"
    };

    private static void EnsureState(PptJob job, PptJobState expected)
    {
        if (job.State != expected) throw new InvalidOperationException($"작업 상태가 올바르지 않습니다: {job.State}");
    }
    private static void EnsureSourceUnchanged(PptJob job)
    {
        if (!File.Exists(job.SourcePath) || File.GetLastWriteTimeUtc(job.SourcePath) != job.SourceLastWriteUtc)
            throw new IOException("계획 수립 후 원본 파일이 변경되었습니다. 다시 분석해 주세요.");
    }
    private static string CreateCopyPath(string source)
    {
        var directory = Path.GetDirectoryName(source)!;
        var name = Path.GetFileNameWithoutExtension(source);
        var extension = Path.GetExtension(source);
        var stem = Path.Combine(directory, $"{name}_BckPPT_{DateTime.Now:yyyyMMdd_HHmmss}");
        var candidate = stem + extension;
        for (var suffix = 2; File.Exists(candidate); suffix++) candidate = $"{stem}_{suffix}{extension}";
        return candidate;
    }
    private static void TryDeleteIncompleteCopy(string path) { try { if (File.Exists(path)) File.Delete(path); } catch { } }

    private static string DescribeDecision(LayoutDecision decision) => decision.Candidate is null
        ? $"• 레이아웃 결정: 신규 생성 · 신뢰도 {decision.Confidence:P0} · {decision.Reason}"
        : $"• 레이아웃 결정: {decision.Kind} · {Path.GetFileName(decision.Candidate.FilePath)} {decision.Candidate.SlideIndex}쪽 · " +
          $"신뢰도 {decision.Confidence:P0} · {decision.Candidate.Reason}";

    private static IReadOnlySet<int> GetSourceSlides(PptTaskPlan plan)
    {
        var slides = new HashSet<int>();
        foreach (var command in plan.Commands)
        {
            if (command.SlideIndex is > 0) slides.Add(command.SlideIndex.Value);
            if (command.SourceSlideIndex is > 0) slides.Add(command.SourceSlideIndex.Value);
        }
        if (slides.Count == 0) slides.Add(1);
        return slides;
    }

    private static IReadOnlySet<int> GetResultSlides(PptTaskPlan plan)
    {
        var slides = new HashSet<int>();
        foreach (var command in plan.Commands)
        {
            if (command.SlideIndex is > 0) slides.Add(command.SlideIndex.Value);
            if (command.InsertAfterSlideIndex is >= 0) slides.Add(command.InsertAfterSlideIndex.Value + 1);
        }
        if (slides.Count == 0) slides.Add(1);
        return slides;
    }
}
