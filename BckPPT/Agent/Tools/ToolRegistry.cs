using System.Text.Json;
using BckPPT.Llm;
using BckPPT.Services;

namespace BckPPT.Agent.Tools;

public sealed record ToolOutcome(string Content, bool IsError);

/// <summary>이름으로 도구를 찾아 실행한다. 모델이 고칠 수 있는 실패는 예외 대신 오류 문장으로 돌려준다.</summary>
public sealed class ToolRegistry
{
    private readonly Dictionary<string, AgentTool> _tools = new(StringComparer.Ordinal);

    public ToolRegistry(IEnumerable<AgentTool> tools)
    {
        foreach (var tool in tools) _tools[tool.Name] = tool;
    }

    public IReadOnlyCollection<AgentTool> Tools => _tools.Values;

    public IReadOnlyList<LlmToolDefinition> Definitions() => _tools.Values.Select(t => t.ToDefinition()).ToList();

    public bool TryGet(string name, out AgentTool tool) => _tools.TryGetValue(name, out tool!);

    public async Task<ToolOutcome> ExecuteAsync(AgentToolContext context, LlmToolCall call, CancellationToken cancellationToken)
    {
        if (!_tools.TryGetValue(call.Name, out var tool))
            return new($"오류: '{call.Name}' 도구는 없습니다. 사용 가능: {string.Join(", ", _tools.Keys)}", true);

        JsonDocument arguments;
        try { arguments = JsonDocument.Parse(string.IsNullOrWhiteSpace(call.ArgumentsJson) ? "{}" : call.ArgumentsJson); }
        catch (JsonException ex) { return new($"오류: 인자가 올바른 JSON이 아닙니다 — {ex.Message}", true); }

        using (arguments)
        {
            if (arguments.RootElement.ValueKind != JsonValueKind.Object)
                return new("오류: 인자는 JSON 객체여야 합니다.", true);
            if (tool.Writes && context.InPlace && ProtectedSlide(context, arguments.RootElement) is int locked)
                return new($"오류: {locked}쪽은 원래 있던 쪽이라 고치지 않습니다(열린 PowerPoint에서 바로 작업 중). " +
                           "새 쪽을 만들어 작업하세요. 사용자가 그 쪽 번호를 적어 요청하면 고칠 수 있습니다.", true);
            try
            {
                return new(await tool.Execute(context, arguments.RootElement, cancellationToken), false);
            }
            catch (OperationCanceledException) { throw; }
            catch (PptUserInterventionException) { throw; }
            catch (AgentToolException ex) { return new($"오류: {ex.Message}", true); }
            catch (Exception ex)
            {
                if (ex.GetBaseException() is PptUserInterventionException intervention) throw intervention;
                // COM 오류(잘못된 슬라이드 번호 등)도 모델이 다음 행동을 고를 수 있게 알린다.
                return new($"오류: {call.Name} 실행 실패 — {ex.GetBaseException().Message}", true);
            }
        }
    }

    /// <summary>
    /// 쓰기 대상 쪽(slide, from_slide~to_slide) 중 이번 작업에서 만들지 않았고 사용자가 요청에 적지도 않은 쪽.
    /// 바로 작업은 원본을 그 자리에서 고치므로, 사용자가 말하지 않은 기존 쪽은 건드리지 않는다.
    /// </summary>
    public static int? ProtectedSlide(AgentToolContext context, JsonElement args)
    {
        var slides = new List<int>();
        if (ToolArgs.OptionalIntOrNull(args, "slide") is int one) slides.Add(one);
        if (ToolArgs.OptionalIntOrNull(args, "from_slide") is int from)
        {
            int to = ToolArgs.OptionalIntOrNull(args, "to_slide") ?? from;
            for (int s = Math.Min(from, to); s <= Math.Max(from, to); s++) slides.Add(s);
        }
        foreach (int slide in slides)
        {
            if (context.Origins.IsCreated(slide)) continue;
            if (context.PermittedOriginalSlides.Contains(context.Origins.OriginalIndexOf(slide))) continue;
            return slide;
        }
        return null;
    }
}
