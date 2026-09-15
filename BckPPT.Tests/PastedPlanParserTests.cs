using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Tests;

public class PastedPlanParserTests
{
    private static IReadOnlyList<ParagraphEntry> Known() =>
    [
        new("S1-12-1", 1, "화살표: 오각형 95", 12, 1, "개방·확장형 보고서"),
        new("S1-12-2", 1, "화살표: 오각형 95", 12, 2, "서식 설계 &"),
        new("S1-12-3", 1, "화살표: 오각형 95", 12, 3, "기존 점검 보고서 등"),
        new("S1-12-4", 1, "화살표: 오각형 95", 12, 4, "구조 분석"),
        new("S1-20-1", 1, "화살표: 오각형 69", 20, 1, "점검·진단 특화 용어·수식 인식 기술"),
        new("S2-31-2", 2, "TextBox 193",      31, 2, "외관조사망도·사진대지 자동 생성 및 정합 소프트웨어"),
    ];

    // ── 기본 형식 ────────────────────────────────────────────────

    [Fact]
    public void 콜론_형식을_읽는다()
    {
        var r = PastedPlanParser.Parse("S1-12-1: 기존 점검·진단 보고서", Known());
        var c = Assert.Single(r.Changes);
        Assert.Equal("S1-12-1", c.Id);
        Assert.Equal("기존 점검·진단 보고서", c.NewText);
        Assert.False(c.Insert);
    }

    [Fact]
    public void 탭과_여러칸_공백을_구분자로_인정한다()
    {
        var r = PastedPlanParser.Parse(
            "S1-12-1\t기존 점검·진단 보고서\nS1-12-2    비정형 정보 구조 분석 및", Known());
        Assert.Equal(2, r.Changes.Count);
        Assert.Equal("기존 점검·진단 보고서", r.Changes[0].NewText);
        Assert.Equal("비정형 정보 구조 분석 및", r.Changes[1].NewText);
    }

    // ── 실제 에이전트 답변 형태 ──────────────────────────────────

    [Fact]
    public void 코드펜스와_설명이_섞여도_ID줄만_집는다()
    {
        const string pasted = """
            좋습니다. 기존 도형은 유지하면서 텍스트만 치환할 부분을 정리하겠습니다.

            ## 1. 치환할 텍스트

            ### [슬라이드 1] 과제추진로드맵

            ```text
            S1-12-1    기존 점검·진단 보고서
            S1-12-2    비정형 정보 구조 분석 및
            S1-12-3    추출·정형화 규격 정의
            ```

            여기서 마지막 문단은 빈 문자열로 치환하는 것을 전제로 했습니다.
            """;

        var r = PastedPlanParser.Parse(pasted, Known());

        Assert.Equal(3, r.Changes.Count);
        Assert.Equal("기존 점검·진단 보고서",   r.Changes[0].NewText);
        Assert.Equal("추출·정형화 규격 정의", r.Changes[2].NewText);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void 마크다운_표_두칸이면_나머지_칸을_쓴다()
    {
        var r = PastedPlanParser.Parse("| S1-12-1 | 기존 점검·진단 보고서 |", Known());
        var c = Assert.Single(r.Changes);
        Assert.Equal("기존 점검·진단 보고서", c.NewText);
    }

    [Fact]
    public void 마크다운_표_여러칸이면_마지막칸을_쓰고_경고한다()
    {
        var r = PastedPlanParser.Parse(
            "| S1-12-1 | 1쪽 | 개방·확장형 보고서 | 기존 점검·진단 보고서 |", Known());
        var c = Assert.Single(r.Changes);
        Assert.Equal("기존 점검·진단 보고서", c.NewText);
        Assert.Single(r.Warnings);
    }

    [Fact]
    public void 화살표_뒷부분을_새_텍스트로_쓴다()
    {
        var r = PastedPlanParser.Parse("S1-12-1 → 기존 점검·진단 보고서", Known());
        Assert.Equal("기존 점검·진단 보고서", Assert.Single(r.Changes).NewText);
    }

    [Fact]
    public void 불릿과_백틱_굵게를_벗겨낸다()
    {
        var r = PastedPlanParser.Parse("- `S1-12-1`: **기존 점검·진단 보고서**", Known());
        Assert.Equal("기존 점검·진단 보고서", Assert.Single(r.Changes).NewText);
    }

    // ── 특수 지시 ────────────────────────────────────────────────

    [Fact]
    public void 빈문단_표시를_빈_문자열로_해석한다()
    {
        var r = PastedPlanParser.Parse("S1-12-4: (빈 문단)", Known());
        Assert.Equal("", Assert.Single(r.Changes).NewText);
    }

    [Fact]
    public void 플러스가_붙으면_문단_추가로_본다()
    {
        var r = PastedPlanParser.Parse("S1-20-1+: 정보 추출·구조화 기술", Known());
        var c = Assert.Single(r.Changes);
        Assert.True(c.Insert);
        Assert.Equal("정보 추출·구조화 기술", c.NewText);
    }

    [Fact]
    public void 이스케이프_줄바꿈은_문단내_줄바꿈으로_바꾼다()
    {
        var r = PastedPlanParser.Parse(@"S1-12-1: 앞줄\n뒷줄", Known());
        Assert.Equal("앞줄\v뒷줄", Assert.Single(r.Changes).NewText);
    }

    // ── 방어 ─────────────────────────────────────────────────────

    [Fact]
    public void 구분자_없는_산문_속_ID_언급은_무시한다()
    {
        var r = PastedPlanParser.Parse(
            "다음은 그대로 유지하는 편이 좋습니다.\n- `S1-12-1` 는 이번 보완과 충돌하지 않습니다.", Known());
        Assert.Empty(r.Changes);
    }

    [Fact]
    public void 같은_문단의_치환과_추가는_서로_다른_작업으로_남는다()
    {
        var r = PastedPlanParser.Parse(
            "S1-20-1: 텍스트·표·도면 등 유형별\nS1-20-1+: 정보 추출·구조화 기술", Known());

        Assert.Equal(2, r.Changes.Count);
        Assert.False(r.Changes[0].Insert);
        Assert.Equal("텍스트·표·도면 등 유형별", r.Changes[0].NewText);
        Assert.True(r.Changes[1].Insert);
        Assert.Equal("정보 추출·구조화 기술", r.Changes[1].NewText);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void 같은_ID가_여러번이면_마지막_값을_쓴다()
    {
        var r = PastedPlanParser.Parse(
            "S1-12-1: 첫 번째 안\n다시 생각해보니\nS1-12-1: 두 번째 안", Known());
        Assert.Equal("두 번째 안", Assert.Single(r.Changes).NewText);
        Assert.Single(r.Warnings);
    }

    [Fact]
    public void 목록에_없는_ID는_경고로_남기고_버린다()
    {
        var r = PastedPlanParser.Parse("S9-99-1: 없는 문단", Known());
        Assert.Empty(r.Changes);
        Assert.Single(r.Warnings);
    }

    [Fact]
    public void 캐리지리턴은_문단을_쪼개지_못하게_공백으로_바꾼다()
    {
        var r = PastedPlanParser.Parse("S1-12-1: 앞\r뒤", Known());
        Assert.DoesNotContain('\r', Assert.Single(r.Changes).NewText);
    }

    [Fact]
    public void 빈_입력은_빈_결과를_준다()
    {
        var r = PastedPlanParser.Parse("   ", Known());
        Assert.Empty(r.Changes);
        Assert.Empty(r.Warnings);
    }

    [Fact]
    public void 소문자_id도_받아들인다()
    {
        var r = PastedPlanParser.Parse("s1-12-1: 기존 점검·진단 보고서", Known());
        Assert.Equal("S1-12-1", Assert.Single(r.Changes).Id);
    }
}
