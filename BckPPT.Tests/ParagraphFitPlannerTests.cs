using BckPPT.Agent;
using BckPPT.Models;

namespace BckPPT.Tests;

public class ParagraphFitPlannerTests
{
    private static List<string> Apply(string[] current, string[] desired)
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([new FakeShape("Body", 0, 0, 100, 100, 50, current)]);
        ppt.SetParagraphTexts("x", ParagraphFitPlanner.Plan(1, "Body", current.Length, desired), markRed: false);
        return ppt.Slides[0][0].Paragraphs;
    }

    [Fact]
    public void 같은_문단_수는_교체만_한다()
    {
        var edits = ParagraphFitPlanner.Plan(1, "Body", 2, ["가", "나"]);
        Assert.All(edits, e => Assert.Equal(ParagraphEditKind.Replace, e.Kind));
        Assert.Equal(["가", "나"], Apply(["a", "b"], ["가", "나"]));
    }

    [Fact]
    public void 늘어나면_마지막_문단_뒤에_순서대로_붙는다()
    {
        Assert.Equal(["1", "2", "3", "4", "5"], Apply(["a", "b"], ["1", "2", "3", "4", "5"]));
    }

    [Fact]
    public void 줄어들면_뒤에서부터_지운다()
    {
        var edits = ParagraphFitPlanner.Plan(1, "Body", 4, ["1"]);
        Assert.Equal(3, edits.Count(e => e.Kind == ParagraphEditKind.Delete));
        Assert.Equal(["1"], Apply(["a", "b", "c", "d"], ["1"]));
    }

    [Fact]
    public void 줄바꿈은_Shift_Enter로_바꾼다()
    {
        Assert.Equal(["첫 줄\v둘째 줄"], Apply(["a"], ["첫 줄\n둘째 줄"]));
    }

    [Fact]
    public void 문단_수는_캐리지리턴으로_센다()
    {
        Assert.Equal(1, ParagraphFitPlanner.CountParagraphs(""));
        Assert.Equal(1, ParagraphFitPlanner.CountParagraphs("한 줄\v같은 문단"));
        Assert.Equal(3, ParagraphFitPlanner.CountParagraphs("a\rb\rc"));
    }

    [Fact]
    public void 빈_목록은_거부한다()
    {
        Assert.Throws<ArgumentException>(() => ParagraphFitPlanner.Plan(1, "Body", 2, []));
    }
}

public class SlideOriginsTests
{
    [Fact]
    public void 복제본은_원본을_기준으로_삼는다()
    {
        var origins = new SlideOrigins();
        origins.RecordDuplicate(source: 33, newIndex: 35);
        Assert.Equal(33, origins.BaselineOf(35));
        Assert.Contains(35, origins.Touched);
    }

    [Fact]
    public void 앞에_끼우면_기록이_함께_밀린다()
    {
        var origins = new SlideOrigins();
        origins.RecordDuplicate(33, 35);      // 35 ← 33
        origins.RecordDuplicate(10, 11);      // 11에 끼움 → 33→34, 35→36
        Assert.Equal(34, origins.BaselineOf(36));
        Assert.Equal(10, origins.BaselineOf(11));
        Assert.Null(origins.BaselineOf(35));
    }

    [Fact]
    public void 작업_사본_번호를_원본_번호로_되돌린다()
    {
        var origins = new SlideOrigins();
        origins.RecordDuplicate(3, 4);    // 원본 3쪽 복제 → 4쪽, 원본 4쪽은 5쪽으로
        origins.RecordDuplicate(1, 1);    // 원본 1쪽 복제 → 맨 앞(1쪽), 나머지 한 칸씩 밀림
        Assert.Equal(1, origins.OriginalIndexOf(1));   // 새 슬라이드 → 복제 원본(원본 1쪽)
        Assert.Equal(1, origins.OriginalIndexOf(2));   // 원본 1쪽
        Assert.Equal(3, origins.OriginalIndexOf(4));   // 원본 3쪽
        Assert.Equal(3, origins.OriginalIndexOf(5));   // 새 슬라이드 → 원본 3쪽
        Assert.Equal(4, origins.OriginalIndexOf(6));   // 원본 4쪽
    }

    [Fact]
    public void 복제본을_다시_복제해도_처음_원본이_기준이다()
    {
        var origins = new SlideOrigins();
        origins.RecordDuplicate(33, 35);
        origins.RecordDuplicate(35, 36);
        Assert.Equal(33, origins.BaselineOf(36));
    }
}

public class PilcrowTests
{
    [Fact]
    public void 미리보기_표기를_흉내_낸_문단은_나눈다()
    {
        Assert.Equal(["가", "나", "다"], ParagraphFitPlanner.SplitPilcrows(["가 ¶ 나 ¶ 다"]));
        Assert.Equal("첫\v둘", ParagraphFitPlanner.NormalizeParagraph("첫 ↵ 둘"));
    }
}
