using BckPPT.Models;
using BckPPT.PowerPoint;

namespace BckPPT.Tests;

/// <summary>2026-10-06 착수보고 작업에서 사람이 잡았던 문제를 규칙이 잡아내는지 본다.</summary>
public class LayoutLintTests
{
    private static readonly LintSettings Settings = new();

    private static TextLineInfo L(int n, string text, LineBreakType type) => new(n, text, type);

    private static LintShape Text(string name, params TextLineInfo[] lines) =>
        new(name, name, 0, 0, 200, 300, lines, BoundHeight: lines.Length * 10f);

    private static LintShape Lines(string name, int count) =>
        Text(name, Enumerable.Range(1, count)
            .Select(i => L(i, "충분히 긴 한 줄", i < count ? LineBreakType.Enter : LineBreakType.None)).ToArray());

    [Fact]
    public void 기준보다_줄이_많으면_잡는다()
    {
        var issues = LayoutLint.Check(new(35, [Lines("C2_rnd", 12)]), new(33, [Lines("C2_rnd", 11)]), Settings);
        var issue = Assert.Single(issues);
        Assert.Equal(LintRule.BaselineLines, issue.Rule);
        Assert.Equal("C2_rnd 12줄 > 기준 11줄 — 1줄 초과, 약 9자 줄이세요", issue.Message);
    }

    [Fact]
    public void 기준과_같거나_적으면_통과한다()
    {
        Assert.Empty(LayoutLint.Check(new(35, [Lines("C2_rnd", 11)]), new(33, [Lines("C2_rnd", 11)]), Settings));
    }

    [Fact]
    public void 자동_줄바꿈으로_넘어간_짧은_꼬리는_고아_단어다()
    {
        var shape = Text("C3_rnd",
            L(1, "첫 문단은 한 줄", LineBreakType.Enter),
            L(2, "AI 기반 공정 최적화 기술", LineBreakType.AutoWrap),
            L(3, "개발", LineBreakType.None));
        var issue = Assert.Single(LayoutLint.Check(new(35, [shape]), null, Settings));
        Assert.Equal(LintRule.OrphanWord, issue.Rule);
        Assert.Equal("C3_rnd 2문단 마지막 줄 '개발'만 남음", issue.Message);
    }

    [Fact]
    public void 의도한_짧은_문단은_고아가_아니다()
    {
        var shape = Text("C3_rnd", L(1, "개요", LineBreakType.Enter), L(2, "긴 설명 문장이 이어집니다", LineBreakType.None));
        Assert.Empty(LayoutLint.Check(new(35, [shape]), null, Settings));
    }

    [Fact]
    public void Shift_Enter_앞의_꼬리도_잡는다()
    {
        var shape = Text("C1_overview",
            L(1, "디지털 트윈 기반 예측 정비 플랫", LineBreakType.AutoWrap),
            L(2, "폼", LineBreakType.ShiftEnter),
            L(3, "둘째 줄", LineBreakType.None));
        var issue = Assert.Single(LayoutLint.Check(new(35, [shape]), null, Settings));
        Assert.Contains("'폼'", issue.Message);
    }

    [Fact]
    public void 제목_둘째_줄이_짧으면_균형을_권한다()
    {
        var title = Text("C2_title", L(1, "스마트 제조 공정 품질 예측 기술", LineBreakType.AutoWrap), L(2, "개발", LineBreakType.None));
        var issue = Assert.Single(LayoutLint.Check(new(35, [title]), null, Settings));
        Assert.Equal(LintRule.TitleBalance, issue.Rule);
        Assert.Equal("C2_title 둘째 줄 '개발' — Shift+Enter로 균형", issue.Message);
    }

    [Fact]
    public void 넘침을_잡는다()
    {
        var shape = Lines("C4_rnd", 3) with { Height = 27f };   // 텍스트 30pt
        var issue = Assert.Single(LayoutLint.Check(new(35, [shape]), null, Settings));
        Assert.Equal(LintRule.Overflow, issue.Rule);
        Assert.Equal("C4_rnd 텍스트가 상자보다 3pt 넘침", issue.Message);
    }

    [Fact]
    public void 같은_역할_상자_중_자간이_다른_것을_잡는다()
    {
        var normal = new TextFormatInfo("맑은 고딕", 11f, 0f, false, "#000000");
        var shapes = new[]
        {
            Lines("C1_overview", 2) with { Format = normal with { Spacing = -1.5f } },
            Lines("C2_overview", 2) with { Format = normal },
            Lines("C3_overview", 2) with { Format = normal },
        };
        var issue = Assert.Single(LayoutLint.Check(new(35, shapes), null, Settings));
        Assert.Equal(LintRule.FormatDrift, issue.Rule);
        Assert.Equal("C1_overview 자간 -1.5 (다른 상자 0)", issue.Message);
    }

    [Fact]
    public void 과반이_없으면_서식_차이를_말하지_않는다()
    {
        var normal = new TextFormatInfo("맑은 고딕", 11f, 0f, false, "#000000");
        var shapes = new[]
        {
            Lines("C1_overview", 2) with { Format = normal with { Spacing = -1.5f } },
            Lines("C2_overview", 2) with { Format = normal },
        };
        Assert.Empty(LayoutLint.Check(new(35, shapes), null, Settings));
    }

    [Fact]
    public void 추가한_도형이_걸쳐_있으면_겹침이다()
    {
        var band = new LintShape("Band", "Band", 0, 100, 500, 30, []);
        var legend = new LintShape("Legend", "Legend", 400, 90, 80, 22, [], IsAdded: true);
        var issue = Assert.Single(LayoutLint.Check(new(35, [band, legend]), null, Settings));
        Assert.Equal(LintRule.Overlap, issue.Rule);
        Assert.Equal("Legend가 Band와 80×12pt 겹침", issue.Message);
    }

    [Fact]
    public void 패널_안에_올려놓은_것은_겹침이_아니다()
    {
        var panel = new LintShape("Panel", "Panel", 0, 0, 500, 300, []);
        var badge = new LintShape("Badge", "Badge", 10, 10, 60, 20, [], IsAdded: true);
        Assert.Empty(LayoutLint.Check(new(35, [panel, badge]), null, Settings));
    }

    [Fact]
    public void 역할은_열_접두어_뒤_이름이다()
    {
        Assert.Equal("rnd", LayoutLint.RoleOf("C2_rnd"));
        Assert.Equal("title", LayoutLint.RoleOf("C10_title"));
        Assert.Null(LayoutLint.RoleOf("Text 19"));
    }
}
