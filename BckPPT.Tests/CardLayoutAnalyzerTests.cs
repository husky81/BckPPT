using BckPPT.Models;
using BckPPT.PowerPoint;

namespace BckPPT.Tests;

/// <summary>착수보고 33쪽 실제 좌표(2026-10-06 COM으로 읽은 값)로 카드 구조 분석을 확인한다.</summary>
public class CardLayoutAnalyzerTests
{
    private static int _id = 1;
    private static ShapeEntry S(string name, float l, float t, float w, float h, string text = "") =>
        new($"@{_id}", name, "Group 171", 1, l, t, w, h, text.Length > 0, text, _id++);

    /// <summary>33쪽 카드 한 장(x 오프셋 dx). 템플릿 도형 이름이 카드마다 같다.</summary>
    private static IEnumerable<ShapeEntry> Card(float dx, string number, string title, string overview)
    {
        yield return S("Shape 20", 262 + dx, 112, 216, 390);
        yield return S("Shape 21", 262 + dx, 112, 216, 69);
        yield return S("Shape 22", 262 + dx, 164, 216, 17);
        yield return S("Shape 23", 271 + dx, 117, 27, 27);
        yield return S("Text 24", 271 + dx, 117, 27, 27, number);
        yield return S("Text 26", 301 + dx, 120, 50, 20, "성능 평가");
        yield return S("Text 27", 272 + dx, 144, 206, 35, title);
        yield return S("Text 15", 269 + dx, 184, 101, 14, "개요");
        yield return S("Text 16", 271 + dx, 198, 204, 42, overview);
        yield return S("Shape 17", 269 + dx, 258, 196, 0);
        yield return S("Text 18", 269 + dx, 261, 130, 14, "문제점");
        yield return S("Text 19", 271 + dx, 274, 204, 42, "해체 시 체계적인 공법 선정 ¶ 순환골재 품질");
        yield return S("Shape 17", 269 + dx, 332, 196, 0);
        yield return S("Text 18", 269 + dx, 335, 130, 14, "연구개발 내용");
        yield return S("Text 19", 271 + dx, 348, 204, 99, "최적 해체 ¶ 전용 장비 ¶ 파쇄 분리");
    }

    private static List<ShapeEntry> Slide33()
    {
        var shapes = new List<ShapeEntry>
        {
            S("사각형: 잘린 한쪽 모서리 407", 32, 93, 896, 422),
            S("TextBox 7", 32, 77, 421, 22, "내역 2. 효율적 시설물 안전관리"),
            S("슬라이드헤더", 29, 11, 835, 46, "4. 지정공모형 후보 아이템_우선순위 도출 결과(6/7)"),
        };
        shapes.AddRange(Card(-220, "3", "시설물안전법 기반 지표", "현행 시설물의 지난 10년간"));
        shapes.AddRange(Card(0, "4", "노후 시설물 해체", "노후 시설물의 안전하고"));
        shapes.AddRange(Card(220, "5", "AI 기반 차량-교량", "AI 기반 교량 구조응답"));
        shapes.AddRange(Card(440, "6", "THM 디지털트윈", "도심지 지반시설물"));
        return shapes;
    }

    [Fact]
    public void 카드_네_장과_역할을_찾는다()
    {
        var result = CardLayoutAnalyzer.Analyze(Slide33());

        Assert.Equal(4, result.Cards.Count);
        Assert.Equal("사각형: 잘린 한쪽 모서리 407", result.Panel!.Name);
        Assert.Equal("TextBox 7", result.Header!.Name);

        var card = result.Cards[1];   // 왼쪽에서 두 번째 = 원래 4번 카드
        Assert.Equal("4", card.Number!.TextPreview);
        Assert.Equal("성능 평가", card.Tag!.TextPreview);
        Assert.Equal("노후 시설물 해체", card.Title!.TextPreview);
        Assert.Equal(["overview", "problem", "rnd"], card.Sections.Select(s => s.Role));
        Assert.Equal("노후 시설물의 안전하고", card.Sections[0].Body.TextPreview);
        Assert.Equal(99f, card.Sections[2].Body.Height);
    }

    [Fact]
    public void 카드는_왼쪽부터_번호를_매긴다()
    {
        var result = CardLayoutAnalyzer.Analyze(Slide33());
        Assert.Equal(["3", "4", "5", "6"], result.Cards.Select(c => c.Number!.TextPreview));
        Assert.True(result.Cards.Zip(result.Cards.Skip(1)).All(p => p.First.Container.Left < p.Second.Container.Left));
    }

    [Theory]
    [InlineData("개요", "overview")]
    [InlineData("문제점", "problem")]
    [InlineData("연구개발 내용", "rnd")]
    [InlineData("추진 일정", "s3")]
    public void 소제목으로_역할_키를_정한다(string label, string role)
    {
        Assert.Equal(role, CardLayoutAnalyzer.RoleOf(label, 2));
    }
}
