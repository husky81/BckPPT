using BckPPT.Application;
using BckPPT.Models;

namespace BckPPT.Tests;

public sealed class OpenFileMatcherTests
{
    private static readonly IReadOnlyList<OpenPresentationInfo> Open =
    [
        new("착수보고_v2.pptx", @"C:\a\착수보고_v2.pptx", 48, true, true),
        new("추진일정_초안.pptx", @"C:\b\추진일정_초안.pptx", 12, false, true),
        new("중간보고_착수참고.pptx", @"C:\c\중간보고_착수참고.pptx", 30, false, true),
        new("프레젠테이션1", "프레젠테이션1", 3, false, false),
    ];

    [Theory]
    [InlineData("추진일정 파일 3쪽 제목 바꿔줘", "추진일정_초안.pptx")]
    [InlineData("추진일정_초안.pptx에서 빈 줄 지워줘", "추진일정_초안.pptx")]
    [InlineData("중간보고 문서에 카드 넣어줘", "중간보고_착수참고.pptx")]
    public void 요청의_파일_이름_조각으로_열린_문서를_찾는다(string request, string expected)
    {
        var match = OpenFileMatcher.Match(request, Open);
        Assert.Equal(FileMatchKind.One, match.Kind);
        Assert.Equal(expected, match.File!.Name);
    }

    [Theory]
    [InlineData("38~41쪽 공백줄 삭제해줘")]
    [InlineData("기획 내용을 정리해서 카드로 만들어줘")]
    [InlineData("이 파일에서 제목만 고쳐줘")]
    [InlineData("프레젠테이션1 파일에 넣어줘")]          // 저장 안 된 문서는 대상이 될 수 없다
    public void 파일을_말하지_않으면_바꾸지_않는다(string request) =>
        Assert.Equal(FileMatchKind.None, OpenFileMatcher.Match(request, Open).Kind);

    [Fact]
    public void 여러_문서가_맞으면_모호하다고_알린다()
    {
        var match = OpenFileMatcher.Match("착수 파일 40쪽 고쳐줘", Open);
        Assert.Equal(FileMatchKind.Ambiguous, match.Kind);
        Assert.Equal(2, match.Candidates.Count);
    }
}
