using BckPPT.Application;

namespace BckPPT.Tests;

public sealed class ConsoleFormattingTests
{
    [Theory]
    [InlineData("remove_blank_lines", """{"from_slide": 38, "to_slide": "41"}""", "빈 줄 정리", "38–41쪽")]
    [InlineData("draw_schedule", """{"style": "gantt", "template_slide": 43, "after": 44}""", "일정표 그리기", "간트형 · 44쪽 뒤")]
    [InlineData("set_paragraphs", """{"slide": 38, "shape": "Group 3/C1_title"}""", "문단 바꾸기", "38쪽 C1_title")]
    [InlineData("duplicate_slide", """{"source": 35, "after": 37}""", "쪽 복제", "35쪽 → 37쪽 뒤")]
    [InlineData("search_source", """{"query": "연구개발 내용"}""", "원문 검색", "'연구개발 내용'")]
    [InlineData("unknown_tool", "{}", "unknown_tool", "")]
    public void 도구를_한글_동작과_인자_요약으로_보인다(string tool, string args, string label, string detail)
    {
        var line = ConsoleFormatting.Describe(tool, args);
        Assert.Equal((label, detail), (line.Label, line.Detail));
    }

    [Fact]
    public void 마무리_도구는_숨기고_깨진_인자도_견딘다()
    {
        Assert.True(ConsoleFormatting.Describe("finish", """{"summary": "끝"}""").Hidden);
        Assert.Equal("도형 살펴보기", ConsoleFormatting.Describe("list_shapes", "{깨짐").Label);
    }

    [Fact]
    public void 결과는_첫_줄만_오류_머리를_떼고_보인다()
    {
        Assert.Equal("38쪽: 빈 줄 6개 삭제", ConsoleFormatting.ResultLine("\n38쪽: 빈 줄 6개 삭제\n39쪽: 빈 줄 10개 삭제", false));
        Assert.Equal("1쪽은 원래 있던 쪽이라 고치지 않습니다", ConsoleFormatting.ResultLine("오류: 1쪽은 원래 있던 쪽이라 고치지 않습니다", true));
        Assert.Equal("완료", ConsoleFormatting.ResultLine("", false));
        Assert.Equal("38쪽 도형 64개", ConsoleFormatting.ResultLine("38쪽 도형 64개 (주소 (이름) | left,top width×height | 텍스트 앞 60자)", false));
    }

    [Fact]
    public void 파일명은_확장자를_남기고_가운데를_줄인다()
    {
        const string name = "착수보고_BckPPT_20261006_224950_BckPPT_20261007_042850.pptx";
        double Width(string s) => s.Length * 7.0;   // 글자당 7px로 본다

        var shortened = ConsoleFormatting.MiddleEllipsis(name, 30 * 7, Width);

        Assert.True(Width(shortened) <= 30 * 7);
        Assert.StartsWith("착수보고_", shortened);
        Assert.EndsWith("042850.pptx", shortened);
        Assert.Contains("…", shortened);
        Assert.Equal(name, ConsoleFormatting.MiddleEllipsis(name, 1000, Width));    // 들어가면 그대로
        Assert.Equal("….pptx", ConsoleFormatting.MiddleEllipsis(name, 6 * 7, Width));
    }

    [Fact]
    public void 폴더는_마지막_단계만_남긴다()
    {
        Assert.Equal(@"…\261006-기존 기획과제…\",
            ConsoleFormatting.ShortenFolder(@"C:\Users\bckim\cellbox\05 KB 기획\261006-기존 기획과제 내용 카드 추가"));
        Assert.Equal(@"Docs\", ConsoleFormatting.ShortenFolder(@"Docs"));
        Assert.Equal("", ConsoleFormatting.ShortenFolder(null));
    }
}
