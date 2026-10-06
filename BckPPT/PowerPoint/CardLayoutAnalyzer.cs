using BckPPT.Models;

namespace BckPPT.PowerPoint;

/// <summary>카드 안의 "소제목 + 본문" 한 쌍. <see cref="Role"/>은 소제목에서 정한 영문 키(overview 등).</summary>
public sealed record CardSection(string Label, string Role, ShapeEntry LabelShape, ShapeEntry Body);

/// <summary>카드 하나. 상자는 모두 Id 주소로 가리킨다(이름이 겹치는 템플릿 대비).</summary>
public sealed record CardLayout(
    int Index,
    ShapeEntry Container,
    ShapeEntry? Number,
    ShapeEntry? Tag,
    ShapeEntry? Title,
    IReadOnlyList<CardSection> Sections,
    IReadOnlyList<ShapeEntry> Members)
{
    public float Right => Container.Left + Container.Width;
}

/// <summary>슬라이드의 카드 묶음. Panel은 카드를 모두 감싸는 바깥 틀, Header는 패널 위 머리말 상자다.</summary>
public sealed record SlideCards(ShapeEntry? Panel, ShapeEntry? Header, IReadOnlyList<CardLayout> Cards);

/// <summary>
/// 카드형 템플릿(착수보고 33쪽처럼 같은 모양 카드가 가로로 놓인 슬라이드)의 구조를 도형 위치만으로 읽는다.
/// 모델이 63개 도형 중 번호·제목·본문을 짝짓다 계속 틀려(2026-10-06 실행 1~9) 이 일을 코드로 옮겼다.
/// <list type="bullet">
/// <item>카드: 텍스트 없는 세로로 긴 사각형</item>
/// <item>번호: 카드 안 정사각형에 가까운 작은 숫자 상자</item>
/// <item>태그: 번호와 같은 높이, 번호 오른쪽의 짧은 상자</item>
/// <item>제목: 첫 소제목보다 위에 있는 나머지 텍스트 중 가장 넓은 상자</item>
/// <item>소제목: 한 줄 높이의 짧은 텍스트. 그 아래 첫 본문 상자와 짝을 짓는다</item>
/// </list>
/// </summary>
public static class CardLayoutAnalyzer
{
    private static readonly (string Keyword, string Role)[] KnownRoles =
    [
        ("개요", "overview"), ("문제", "problem"), ("한계", "problem"), ("필요", "need"),
        ("연구개발", "rnd"), ("내용", "rnd"), ("목표", "goal"), ("성과", "outcome"), ("기대", "effect"),
    ];

    public static string RoleOf(string label, int index)
    {
        foreach (var (keyword, role) in KnownRoles)
            if (label.Replace(" ", "").Contains(keyword)) return role;
        return $"s{index + 1}";
    }

    public static SlideCards Analyze(IReadOnlyList<ShapeEntry> shapes)
    {
        float widest = shapes.Count == 0 ? 0 : shapes.Max(s => s.Width);
        var cards = shapes
            .Where(s => !s.HasText && s.Height >= 150 && s.Width >= 80 && s.Width < widest * 0.6)
            .OrderBy(s => s.Left).ThenBy(s => s.Top)
            .ToList();
        // 카드끼리 겹치면(카드 안 장식 사각형 등) 큰 것만 카드로 본다.
        cards = cards.Where(c => !cards.Any(o => !ReferenceEquals(o, c) && Inside(c, o) && Area(o) > Area(c))).ToList();

        var layouts = new List<CardLayout>();
        for (int i = 0; i < cards.Count; i++)
        {
            var card = cards[i];
            var members = shapes.Where(s => !ReferenceEquals(s, card) && !cards.Contains(s) && CenterIn(s, card)).ToList();
            layouts.Add(AnalyzeCard(i + 1, card, members));
        }

        ShapeEntry? panel = cards.Count == 0 ? null : shapes
            .Where(s => !s.HasText && !cards.Contains(s) && cards.All(c => CenterIn(c, s)))
            .OrderBy(Area).FirstOrDefault();
        ShapeEntry? header = panel is null ? null : shapes
            .Where(s => s.HasText && s.Top < panel.Top && s.Top + s.Height <= panel.Top + 30 && s.Top > panel.Top - 60 &&
                        s.Left >= panel.Left - 5 && s.Left < panel.Left + panel.Width / 2)
            .OrderByDescending(s => s.Top).FirstOrDefault();
        return new SlideCards(panel, header, layouts);
    }

    private static CardLayout AnalyzeCard(int index, ShapeEntry card, List<ShapeEntry> members)
    {
        var texts = members.Where(s => s.HasText).OrderBy(s => s.Top).ThenBy(s => s.Left).ToList();
        var number = texts.FirstOrDefault(s => s.Width <= 45 && Math.Abs(s.Width - s.Height) <= 8 &&
                                               s.TextPreview.Trim().All(char.IsDigit) && s.TextPreview.Trim().Length > 0);
        var tag = number is null ? null : texts.FirstOrDefault(s => !ReferenceEquals(s, number) &&
            Math.Abs(s.Top - number.Top) <= 12 && s.Left > number.Left && s.Height <= 30);

        bool IsLabel(ShapeEntry s) => !ReferenceEquals(s, number) && !ReferenceEquals(s, tag) &&
                                      s.Height <= 22 && s.TextPreview.Trim().Length is > 0 and <= 12 && !s.TextPreview.Contains('¶');
        var labels = texts.Where(IsLabel).ToList();
        float firstLabelTop = labels.Count > 0 ? labels.Min(l => l.Top) : float.MaxValue;

        var title = texts
            .Where(s => !ReferenceEquals(s, number) && !ReferenceEquals(s, tag) && !labels.Contains(s) && s.Top < firstLabelTop)
            .OrderByDescending(s => s.Width).ThenBy(s => s.Top).FirstOrDefault();

        var sections = new List<CardSection>();
        for (int i = 0; i < labels.Count; i++)
        {
            var label = labels[i];
            float nextTop = i + 1 < labels.Count ? labels[i + 1].Top : float.MaxValue;
            var body = texts.FirstOrDefault(s => !labels.Contains(s) && !ReferenceEquals(s, title) &&
                                                 s.Top > label.Top && s.Top < nextTop);
            if (body is null) continue;
            string text = label.TextPreview.Trim();
            sections.Add(new CardSection(text, RoleOf(text, i), label, body));
        }
        return new CardLayout(index, card, number, tag, title, sections, members);
    }

    /// <summary>표시용 요약. 모델이나 사람이 분석 결과를 확인할 때 쓴다.</summary>
    public static string Describe(SlideCards slide)
    {
        var lines = new List<string> { $"카드 {slide.Cards.Count}개, 패널 {Addr(slide.Panel)}, 머리말 {Addr(slide.Header)}" };
        foreach (var c in slide.Cards)
            lines.Add($"카드 {c.Index}: 번호 {Addr(c.Number)}, 태그 {Addr(c.Tag)}, 제목 {Addr(c.Title)}, " +
                      string.Join(", ", c.Sections.Select(s => $"{s.Label}({s.Role}) {Addr(s.Body)}")));
        return string.Join("\n", lines);
    }

    public static string Addr(ShapeEntry? s) => s is null ? "없음" : s.Id > 0 ? $"@{s.Id}" : s.Address;

    private static bool CenterIn(ShapeEntry s, ShapeEntry box)
    {
        float cx = s.Left + s.Width / 2, cy = s.Top + s.Height / 2;
        return cx >= box.Left && cx <= box.Left + box.Width && cy >= box.Top && cy <= box.Top + box.Height;
    }

    private static bool Inside(ShapeEntry inner, ShapeEntry outer) =>
        inner.Left >= outer.Left - 1 && inner.Top >= outer.Top - 1 &&
        inner.Left + inner.Width <= outer.Left + outer.Width + 1 && inner.Top + inner.Height <= outer.Top + outer.Height + 1;

    private static float Area(ShapeEntry s) => s.Width * s.Height;
}
