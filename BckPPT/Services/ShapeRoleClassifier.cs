using BckPPT.Models;

namespace BckPPT.Services;

/// <summary>규칙 기반 Shape 역할 분류기.</summary>
public static class ShapeRoleClassifier
{
    /// <summary>전체 프레젠테이션 분석 결과를 기반으로 각 shape의 역할을 분류.</summary>
    public static Dictionary<(int Slide, string Shape), ShapeRole> Classify(
        PresentationAnalysis analysis,
        Dictionary<(string Prs, int Slide, string Shape), ShapeRoleLabel> userLabels,
        string presentationName)
    {
        var result = new Dictionary<(int, string), ShapeRole>();

        // ── 1단계: MasterTitle 탐지 (동일 텍스트 + 유사 위치 3+ 슬라이드 반복) ──
        var textPositionCounts = new Dictionary<string, int>();
        foreach (var shape in analysis.Shapes)
        {
            if (shape.IsEmpty || shape.IsImage) continue;
            var key = NormalizeText(shape.FullText);
            if (key.Length > 0 && key.Length < 100)
                textPositionCounts[key] = textPositionCounts.GetValueOrDefault(key) + 1;
        }
        var masterTexts = textPositionCounts
            .Where(kv => kv.Value >= 3)
            .Select(kv => kv.Key)
            .ToHashSet();

        // ── 2단계: 슬라이드별 최대 폰트 크기 ──
        var slideMaxFont = new Dictionary<int, float>();
        foreach (var shape in analysis.Shapes)
        {
            if (shape.IsEmpty || shape.IsImage || shape.FontSize <= 0) continue;
            if (!slideMaxFont.TryGetValue(shape.SlideIndex, out var max) || shape.FontSize > max)
                slideMaxFont[shape.SlideIndex] = shape.FontSize;
        }

        // ── 3단계: 각 shape 분류 ──
        foreach (var shape in analysis.Shapes)
        {
            var key = (shape.SlideIndex, shape.ShapeName);

            // 사용자 라벨 최우선
            if (userLabels.TryGetValue((presentationName, shape.SlideIndex, shape.ShapeName), out var label))
            {
                result[key] = label.Role;
                continue;
            }

            if (shape.IsImage)
            {
                result[key] = ClassifyImage(shape);
                continue;
            }

            if (shape.IsEmpty)
            {
                result[key] = ShapeRole.Unknown;
                continue;
            }

            result[key] = ClassifyText(shape, masterTexts, slideMaxFont);
        }

        return result;
    }

    private static ShapeRole ClassifyText(PptSlideInfo shape,
        HashSet<string> masterTexts, Dictionary<int, float> slideMaxFont)
    {
        var name = shape.ShapeName;
        var text = shape.FullText.Trim();

        // Shape 이름 패턴
        if (name.Contains("Slide Number", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("슬라이드 번호", StringComparison.Ordinal))
            return ShapeRole.PageNumber;

        if (name.Contains("Footer", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("바닥글", StringComparison.Ordinal))
            return ShapeRole.Footer;

        // 페이지 번호: 매우 짧은 텍스트 (숫자만)
        if (text.Length <= 4 && int.TryParse(text.Trim(), out _))
            return ShapeRole.PageNumber;

        // MasterTitle
        if (masterTexts.Contains(NormalizeText(text)))
            return ShapeRole.Chapter;

        // Title (Shape 이름 기반)
        if (name.Contains("Title", StringComparison.OrdinalIgnoreCase) ||
            name.Contains("제목", StringComparison.Ordinal))
            return ShapeRole.Title;

        // 최대 폰트 = SlideTitle
        if (shape.FontSize > 0 &&
            slideMaxFont.TryGetValue(shape.SlideIndex, out var maxFs) &&
            Math.Abs(shape.FontSize - maxFs) < 0.5f &&
            shape.LineCount <= 3)
            return ShapeRole.Title;

        // TODO: 불릿 기반 ListItem 분류 (TextLineLayout 정보 필요, Phase 2에서 개선)

        return ShapeRole.Body;
    }

    private static ShapeRole ClassifyImage(PptSlideInfo shape)
    {
        // 매우 작은 이미지 → BulletIcon 후보
        if (shape.BoxHeight < 30 && shape.BoundHeight < 30)
            return ShapeRole.BulletIcon;

        return ShapeRole.ContentImage;
    }

    private static string NormalizeText(string text)
        => text.Trim().Replace("\r", "").Replace("\n", "").Replace(" ", "");
}
