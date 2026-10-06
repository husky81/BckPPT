using System.Text.RegularExpressions;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.PowerPoint;

public enum LintRule { BaselineLines, OrphanWord, Overflow, TitleBalance, FormatDrift, Overlap }

public sealed record LintIssue(LintRule Rule, string Address, string Message);

/// <summary>판정에 필요한 도형 한 개의 측정값. 텍스트가 없는 도형은 겹침 판정에만 쓴다.</summary>
public sealed record LintShape(
    string                     Address,
    string                     Name,
    float                      Left,
    float                      Top,
    float                      Width,
    float                      Height,
    IReadOnlyList<TextLineInfo> Lines,
    float                      BoundHeight = 0f,
    TextFormatInfo?            Format      = null,
    bool                       IsAdded     = false,
    int                        Id          = 0)
{
    public bool HasText => Lines.Count > 0;

    /// <summary>기준 슬라이드의 같은 도형. 복제해도 Id는 유지되므로, 이름을 바꾼 뒤에도 Id로 짝을 짓는다.</summary>
    public LintShape? Counterpart(LintSlide? baseline) =>
        baseline?.Shapes.FirstOrDefault(b => Id > 0 && b.Id == Id)
        ?? baseline?.Shapes.FirstOrDefault(b => b.Address == Address);
}

public sealed record LintSlide(int SlideIndex, IReadOnlyList<LintShape> Shapes);

public sealed record LintSettings(
    int                  OrphanMaxChars          = 4,
    float                TitleSecondLineMinRatio = 0.3f,
    float                OverflowTolerance       = 1f,
    float                OverlapTolerance        = 1f,
    IReadOnlyList<string>? TitleRoles            = null)
{
    public IReadOnlyList<string> TitleRolesOrDefault => TitleRoles ?? ["title", "제목"];

    public static LintSettings FromRules(PptRules rules) => new(
        rules.OrphanMaxChars, rules.TitleSecondLineMinRatio, rules.OverflowTolerance,
        TitleRoles: rules.TitleRoles);
}

/// <summary>
/// 모델의 눈대중 대신 코드가 레이아웃을 판정한다. 입력은 PowerPoint가 실제로 나눈 줄과 상자 크기이고,
/// 출력은 모델이 바로 고칠 수 있는 문장이다. COM에 의존하지 않으므로 측정값만으로 테스트한다.
/// </summary>
public static class LayoutLint
{
    // "C2_rnd", "C10_title"처럼 열 번호와 역할을 붙인 이름. 같은 역할끼리 서식을 비교한다.
    private static readonly Regex RoleName = new(@"^[A-Za-z]+\d+_(.+)$");

    public static IReadOnlyList<LintIssue> Check(LintSlide slide, LintSlide? baseline, LintSettings settings)
    {
        var issues = new List<LintIssue>();
        foreach (var shape in slide.Shapes.Where(s => s.HasText && !Unchanged(s, baseline)))
        {
            if (shape.Counterpart(baseline) is { HasText: true } basis &&
                shape.Lines.Count > basis.Lines.Count)
                issues.Add(new(LintRule.BaselineLines, shape.Address,
                    $"{shape.Address} {shape.Lines.Count}줄 > 기준 {basis.Lines.Count}줄{ExcessHint(shape.Lines, basis.Lines.Count)}"));

            if (IsTitle(shape, settings)) CheckTitleBalance(shape, settings, issues);
            else CheckOrphans(shape, settings, issues);

            if (shape.BoundHeight > shape.Height + settings.OverflowTolerance && !TemplateOverflowsAlike(shape, baseline, settings))
                issues.Add(new(LintRule.Overflow, shape.Address,
                    $"{shape.Address} 텍스트가 상자보다 {shape.BoundHeight - shape.Height:0.#}pt 넘침{WidenHint(shape)}"));
        }
        CheckFormatDrift(slide, baseline, issues);
        CheckOverlaps(slide, settings, issues);
        return issues;
    }

    /// <summary>
    /// 기준과 주소·줄·크기가 모두 같은 상자는 이번 작업에서 손대지 않은 것이다.
    /// 템플릿에 원래 있던 문제까지 고치라고 하면 모델이 요청 밖의 상자를 건드린다(2026-10-06 실측).
    /// </summary>
    public static bool Unchanged(LintShape shape, LintSlide? baseline) =>
        shape.Counterpart(baseline) is { } basis &&
        basis.Height == shape.Height && basis.Width == shape.Width &&
        basis.Lines.Select(l => l.Text).SequenceEqual(shape.Lines.Select(l => l.Text));

    /// <summary>
    /// 기준 줄 수를 넘을 때 얼마나 줄여야 하는지. 기준을 넘는 줄의 글자 수 합이 대략 덜어낼 양이다.
    /// "3줄 > 기준 2줄"만으로는 모델이 감을 못 잡고 문구를 거의 그대로 다시 넣었다(2026-10-06 실측).
    /// </summary>
    public static string ExcessHint(IReadOnlyList<TextLineInfo> lines, int budget)
    {
        if (lines.Count <= budget) return "";
        int chars = lines.Skip(budget).Sum(l => l.Text.Trim().Length);
        return $" — {lines.Count - budget}줄 초과, 약 {chars}자 줄이세요";
    }

    public static string Describe(IReadOnlyList<LintIssue> issues) => issues.Count == 0
        ? "위반 없음"
        : $"위반 {issues.Count}건\n" + string.Join("\n", issues.Select(i => $"- [{i.Rule}] {i.Message}"));

    /// <summary>"C2_rnd" → "rnd". 규약을 따르지 않는 이름은 null.</summary>
    public static string? RoleOf(string name)
    {
        var match = RoleName.Match(name);
        return match.Success ? match.Groups[1].Value : null;
    }

    /// <summary>
    /// 줄 목록을 "보이는 줄 묶음"으로 나눈다. Enter(문단)와 Shift+Enter 모두 묶음의 끝이다.
    /// 자동 줄바꿈만 묶음 안에 남으므로, 묶음의 마지막 줄이 짧으면 넘어간 꼬리다.
    /// </summary>
    public static IReadOnlyList<IReadOnlyList<TextLineInfo>> Segments(IReadOnlyList<TextLineInfo> lines)
    {
        var segments = new List<IReadOnlyList<TextLineInfo>>();
        var current = new List<TextLineInfo>();
        foreach (var line in lines)
        {
            current.Add(line);
            if (line.BreakType == LineBreakType.AutoWrap) continue;
            segments.Add(current);
            current = [];
        }
        if (current.Count > 0) segments.Add(current);
        return segments;
    }

    private static bool IsTitle(LintShape shape, LintSettings settings)
    {
        var role = RoleOf(shape.Name) ?? shape.Name;
        return settings.TitleRolesOrDefault.Any(t => role.Contains(t, StringComparison.OrdinalIgnoreCase));
    }

    private static void CheckOrphans(LintShape shape, LintSettings settings, List<LintIssue> issues)
    {
        int paragraph = 1;
        foreach (var segment in Segments(shape.Lines))
        {
            var last = segment[^1];
            if (segment.Count >= 2 && VisibleLength(last.Text) is > 0 and var length && length <= settings.OrphanMaxChars)
                issues.Add(new(LintRule.OrphanWord, shape.Address,
                    $"{shape.Address} {paragraph}문단 마지막 줄 '{last.Text.Trim()}'만 남음{WidenHint(shape)}"));
            if (last.BreakType != LineBreakType.ShiftEnter) paragraph++;
        }
    }

    private static void CheckTitleBalance(LintShape shape, LintSettings settings, List<LintIssue> issues)
    {
        foreach (var segment in Segments(shape.Lines))
        {
            if (segment.Count != 2) continue;
            int first = VisibleLength(segment[0].Text), second = VisibleLength(segment[1].Text);
            if (first > 0 && second < first * settings.TitleSecondLineMinRatio)
                issues.Add(new(LintRule.TitleBalance, shape.Address,
                    $"{shape.Address} 둘째 줄 '{segment[1].Text.Trim()}' — Shift+Enter로 균형"));
        }
    }

    private static void CheckFormatDrift(LintSlide slide, LintSlide? baseline, List<LintIssue> issues)
    {
        var groups = slide.Shapes
            .Where(s => s.Format is not null && RoleOf(s.Name) is not null)
            .GroupBy(s => RoleOf(s.Name)!)
            .Where(g => g.Count() >= 2);
        foreach (var group in groups)
        {
            Compare(group, f => f.FontName, "글꼴", v => v);
            Compare(group, f => f.Size, "크기", v => $"{v:0.#}");
            Compare(group, f => f.Spacing, "자간", v => $"{v:0.#}");
        }

        void Compare<T>(IEnumerable<LintShape> shapes, Func<TextFormatInfo, T> pick, string label, Func<T, string> show)
            where T : notnull
        {
            var counts = shapes.GroupBy(s => pick(s.Format!)).OrderByDescending(g => g.Count()).ToList();
            // 과반이 없으면 어느 쪽이 이탈인지 알 수 없으므로 말하지 않는다.
            if (counts.Count < 2 || counts[0].Count() == counts[1].Count()) return;
            var majority = counts[0].Key;
            // 비교는 모든 상자로 하되, 지적은 이번에 바뀐 상자만 한다.
            foreach (var odd in counts.Skip(1).SelectMany(g => g).Where(s => !Unchanged(s, baseline)))
                issues.Add(new(LintRule.FormatDrift, odd.Address,
                    $"{odd.Address} {label} {show(pick(odd.Format!))} (다른 상자 {show(majority)})"));
        }
    }

    private static void CheckOverlaps(LintSlide slide, LintSettings settings, List<LintIssue> issues)
    {
        foreach (var added in slide.Shapes.Where(s => s.IsAdded))
        {
            foreach (var other in slide.Shapes)
            {
                if (ReferenceEquals(other, added) || other.Address == added.Address) continue;
                float w = Math.Min(added.Left + added.Width, other.Left + other.Width) - Math.Max(added.Left, other.Left);
                float h = Math.Min(added.Top + added.Height, other.Top + other.Height) - Math.Max(added.Top, other.Top);
                if (w <= settings.OverlapTolerance || h <= settings.OverlapTolerance) continue;
                // 패널·배경 위에 올려놓은 것은 의도한 배치다. 걸쳐 있을 때만 문제로 본다.
                if (Contains(other, added, settings.OverlapTolerance)) continue;
                issues.Add(new(LintRule.Overlap, added.Address,
                    $"{added.Address}가 {other.Address}와 {w:0.#}×{h:0.#}pt 겹침"));
            }
        }
    }

    /// <summary>
    /// 기준 상자도 원래 넘쳐 있고(카드 배경 안에서 넘치게 디자인한 템플릿) 줄 수가 기준 이내면 넘침으로 보지 않는다.
    /// 착수보고 33쪽 본문 상자가 모두 이랬는데, 같은 줄 수인데도 지적해 루프가 헛돌았다(2026-10-06 실행 10).
    /// </summary>
    private static bool TemplateOverflowsAlike(LintShape shape, LintSlide? baseline, LintSettings settings) =>
        shape.Counterpart(baseline) is { HasText: true } basis &&
        basis.BoundHeight > basis.Height + settings.OverflowTolerance &&
        shape.Lines.Count <= basis.Lines.Count;

    /// <summary>
    /// 태그·배지처럼 짧은 한 문단 상자가 줄바꿈되면 문구가 아니라 폭이 문제다.
    /// 모델이 "시공 안전성 ● 유지관리 효율 ●"를 50pt 상자에 넣고 문구만 고치며 턴을 썼다(2026-10-06 실측).
    /// </summary>
    private static string WidenHint(LintShape shape)
    {
        bool shortSingle = Segments(shape.Lines).Count(s => s[^1].BreakType != LineBreakType.ShiftEnter) == 1 &&
                           shape.Lines.Sum(l => l.Text.Trim().Length) <= 30;
        return shortSingle ? $" — 짧은 한 줄 상자면 문구 대신 set_geometry로 폭을 넓히세요(지금 {shape.Width:0.#}pt)" : "";
    }

    private static bool Contains(LintShape outer, LintShape inner, float tolerance) =>
        inner.Left >= outer.Left - tolerance && inner.Top >= outer.Top - tolerance &&
        inner.Left + inner.Width <= outer.Left + outer.Width + tolerance &&
        inner.Top + inner.Height <= outer.Top + outer.Height + tolerance;

    private static int VisibleLength(string text) => text.Count(c => !char.IsWhiteSpace(c));
}

/// <summary>COM으로 슬라이드의 측정값을 모아 <see cref="LintSlide"/>를 만든다.</summary>
public static class SlideLintCollector
{
    public static LintSlide Collect(
        IPptComService ppt, string path, int slideIndex, IReadOnlySet<string>? addedAddresses = null)
    {
        var shapes = new List<LintShape>();
        foreach (var entry in ppt.ListShapes(path, slideIndex))
        {
            IReadOnlyList<TextLineInfo> lines = [];
            float bound = 0f;
            TextFormatInfo? format = null;
            if (entry.HasText)
            {
                lines = ppt.GetTextLineBreaks(path, slideIndex, entry.Address);
                bound = ppt.GetShapeInfo(path, slideIndex, entry.Address)?.BoundHeight ?? 0f;
                try { format = ppt.GetTextFormat(path, slideIndex, entry.Address); } catch { }
            }
            shapes.Add(new LintShape(entry.Address, entry.Name, entry.Left, entry.Top, entry.Width, entry.Height,
                lines, bound, format, addedAddresses?.Contains(entry.Address) == true || addedAddresses?.Contains(entry.Name) == true,
                entry.Id));
        }
        return new LintSlide(slideIndex, shapes);
    }
}
