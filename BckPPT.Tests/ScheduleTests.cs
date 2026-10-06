using System.IO;
using System.Text.Json;
using BckPPT.Agent.Tools;
using BckPPT.Models;
using BckPPT.Schedule;

namespace BckPPT.Tests;

public sealed class ScheduleTests
{
    // 착수보고 43쪽에서 뽑은 md의 축약본.
    private const string Md = """
        # 2. 추진 일정(안)

        - 기간: 2026-07 ~ 2027-07

        ## 1. 대내외 환경분석

        | 항목 | 시작 | 끝 |
        |---|---|---|
        | 동향조사 | 2026-07-01 | 2026-07-31 |
        | 통계(CSI)분석 및 현장의견 수렴 | 2026-08-01 | 2026-09-15 |

        ## 3. 단기 상세기획

        ### 3-1. 품목지정

        | 항목 | 시작 | 끝 |
        |---|---|---|
        | 전략계획서 및 RFP 마련 | 2026-11-01 | 2026-12-31 |

        ### 3-2. 지정공모

        | 항목 | 시작 | 끝 |
        |---|---|---|
        | 세부 연구내용 작성 | 2027-01-01 | 2027-02-28 |
        | 예산 심의 의견 반영 | 2027-06-01 | 2027-07-31 |

        ## 운영: 보고회

        | 보고 | 시점 |
        |---|---|
        | 착수보고 | 2026-10-01 |
        | 중간보고 | 2026-12 하순 |
        """;

    [Fact]
    public void 일정_md를_구분_하위구분_마일스톤으로_읽는다()
    {
        var doc = ScheduleMarkdown.Parse(Md);

        Assert.Equal("2. 추진 일정(안)", doc.Title);
        Assert.Equal(["대내외 환경분석", "단기 상세기획"], doc.Groups.Select(g => g.Name));
        Assert.False(doc.Groups[0].HasSubsections);
        Assert.Equal(["품목지정", "지정공모"], doc.Groups[1].Sections.Select(s => s.Name));
        Assert.Equal(new DateTime(2026, 9, 15), doc.Groups[0].Sections[0].Tasks[1].End);
        var report = Assert.Single(doc.MilestoneGroups);
        Assert.Equal(("운영", "보고회"), (report.Label, report.Name));
        Assert.Equal(new DateTime(2026, 12, 25), report.Milestones[1].Date);
        Assert.Equal(new DateTime(2026, 7, 1), doc.First);
        Assert.Equal(new DateTime(2027, 7, 31), doc.Last);
    }

    [Theory]
    [InlineData("2026.10.16", false, "2026-10-16")]
    [InlineData("2026-12", false, "2026-12-01")]
    [InlineData("2026-12", true, "2026-12-31")]
    [InlineData("2026-12 상순", true, "2026-12-10")]
    [InlineData("2026-12 하순", false, "2026-12-21")]
    public void 날짜는_일_월_순까지_읽는다(string text, bool end, string expected)
    {
        Assert.True(ScheduleMarkdown.TryParseDate(text, end, out var date));
        Assert.Equal(expected, date.ToString("yyyy-MM-dd"));
    }

    [Fact]
    public void 기간이_길면_분기_눈금을_쓴다()
    {
        Assert.Equal(13, ScheduleRenderer.Periods(new(2026, 7, 1), new(2027, 7, 31)).Count);
        var quarters = ScheduleRenderer.Periods(new(2026, 1, 1), new(2028, 12, 31));
        Assert.Equal(12, quarters.Count);
        Assert.Equal("1Q", quarters[0].Label);
    }

    [Fact]
    public void 표형_일정표는_표를_만들고_막대와_눈금을_한_그룹으로_묶는다()
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([]);
        var doc = ScheduleMarkdown.Parse(Md);

        var result = new ScheduleRenderer(ppt, "work.pptx").Draw(1, doc);

        var table = Assert.IsType<TableRequest>(ppt.Slides[0].Single(s => s.Name == result.TableName).Spec);
        int cols = table.Cells[0].Count;
        Assert.Equal(4, cols);                                                // 단계·하위·항목·기간(월 눈금은 표 밖 도형)
        Assert.Equal("세부 추진 항목", table.Cells[0][2]);
        Assert.Equal(1, table.HeaderRows);                                    // 머리 2행이 모두 병합되면 PowerPoint가 행을 없앤다
        Assert.Equal("통계(CSI)분석 및 현장의견 수렴", table.Cells[2][2]);
        Assert.Equal("8.1~9.15", table.Cells[2][3]);
        Assert.Equal("품목지정", table.Cells[3][1]);                           // 하위 구분 칸(한 줄짜리는 병합 없음)
        Assert.Contains(new TableMerge(5, 2, 6, 2), table.Merges!);            // 지정공모 하위 칸(두 줄)
        Assert.Contains(new TableMerge(2, 1, 3, 2), table.Merges!);            // 하위 없는 구분은 하위 칸까지 합친다

        var overlay = ppt.Slides[0].Where(s => s.Spec is TextShapeRequest).ToList();
        var specs = overlay.Select(s => (TextShapeRequest)s.Spec!).ToList();
        Assert.Equal(5, specs.Count(r => r.Kind == TextShapeKind.RoundedRectangle));       // 막대 5
        Assert.Equal(2, specs.Count(r => r.Kind == TextShapeKind.Diamond));                // ◆ 2
        Assert.Equal(["착수보고", "중간보고"], specs.Where(r => r.Kind == TextShapeKind.Rectangle && r.AutoWidth).Select(r => r.Text));
        Assert.Equal(["2026", "2027"], specs.Where(r => r.Text.StartsWith("20")).Select(r => r.Text));       // 연 머리글
        Assert.Equal(13, specs.Count(r => r.Text.EndsWith('월')));                                          // 월 머리글
        Assert.Null(table.BorderColor);                                                    // 칸 테두리 대신 선 도형
        Assert.Contains(specs, r => r.LineColor is not null && r.FillColor is null);       // 바깥 테두리
        Assert.All(overlay, s => Assert.Equal(result.Name, s.Group));                      // 막대·눈금은 한 그룹(표는 묶을 수 없다)

        // 첫 막대는 7월 칸 왼쪽 끝에서 시작해 그 칸 끝에서 끝난다.
        var july = specs.First(r => r.Text == "7월");                                     // 2026년 7월
        var bar = overlay.First(s => ((TextShapeRequest)s.Spec!).Kind == TextShapeKind.RoundedRectangle);
        Assert.Equal(table.Left + table.ColumnWidths!.Sum(), july.Left, 1);                   // 눈금은 표 바로 오른쪽
        Assert.Equal(july.Left, bar.Left, 1);
        Assert.Equal(july.Width, bar.Width, 1);
    }

    [Fact]
    public void 간트형은_막대_옆에_이름을_적는다()
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([]);
        new ScheduleRenderer(ppt, "work.pptx").Draw(1, ScheduleMarkdown.Parse(Md), new ScheduleOptions(ScheduleStyle.Gantt));

        var labels = ppt.Slides[0].Select(s => s.Spec).OfType<TextShapeRequest>().Where(r => r.Text.Length > 0).Select(r => r.Text).ToList();
        Assert.Contains("동향조사", labels);
        // 끝 칸까지 가는 막대는 이름을 왼쪽에 둔다.
        var last = ppt.Slides[0].Select(s => s.Spec).OfType<TextShapeRequest>().Single(r => r.Text == "예산 심의 의견 반영");
        Assert.NotNull(last.RightEdge);
    }

    [Fact]
    public void 막대_안_이름형은_들어가면_막대_안에_안_들어가면_옆에_쓴다()
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([]);
        new ScheduleRenderer(ppt, "work.pptx").Draw(1, ScheduleMarkdown.Parse(Md), new ScheduleOptions(ScheduleStyle.InBar));

        var requests = ppt.Slides[0].Select(s => s.Spec).OfType<TextShapeRequest>().ToList();
        var inside = requests.Single(r => r.Text == "세부 연구내용 작성");                       // 두 달짜리 막대
        Assert.Equal(TextShapeKind.RoundedRectangle, inside.Kind);
        Assert.Equal(TextAlign.Center, inside.Align);
        Assert.Equal("#FFFFFF", inside.FontColor);
        var outside = requests.Single(r => r.Text == "동향조사");                               // 한 달짜리 막대: 7월 칸이 좁으면 옆에
        Assert.True(outside.Kind == TextShapeKind.RoundedRectangle || outside.AutoWidth);
    }

    [Fact]
    public async Task 일정표_도구는_빈_새_쪽을_만들어_그린다()
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([new FakeShape("슬라이드헤더", 29, 11, 835, 46, 40, "2. 추진 일정(안)"), new FakeShape("셀", 33, 80, 100, 20, 10, "구분")]);
        var dir = Path.Combine(Path.GetTempPath(), "bckppt-sch-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(dir);
        var mdPath = Path.Combine(dir, "일정.md");
        File.WriteAllText(mdPath, Md);
        try
        {
            var ctx = new AgentToolContext(ppt, Path.Combine(dir, "work.pptx"));
            var tool = PptAgentTools.Create(includeSources: false).Single(t => t.Name == "draw_schedule");
            var args = JsonDocument.Parse(JsonSerializer.Serialize(new { md_path = mdPath, template_slide = 1, after = 1, title = "2. 추진 일정(안)_세부" })).RootElement;

            var result = await tool.Execute(ctx, args, CancellationToken.None);

            Assert.Contains("2쪽에 그렸습니다", result);
            Assert.Equal(2, ppt.Slides.Count);
            Assert.DoesNotContain(ppt.Slides[1], s => s.Name == "셀");                           // 비웠다
            Assert.Equal(["2. 추진 일정(안)_세부"], ppt.Slides[1].Single(s => s.Name == "슬라이드헤더").Paragraphs);
            Assert.Contains(ppt.Slides[1], s => s.Spec is TableRequest);
        }
        finally { Directory.Delete(dir, true); }
    }

    [Fact]
    public void 그룹_이름으로_지우면_그룹을_통째로_지운다()
    {
        var ppt = new FakePptComService();
        ppt.Slides.Add([new FakeShape("a", 0, 0, 1, 1, 1) { Group = "G" }, new FakeShape("b", 0, 0, 1, 1, 1) { Group = "G" }, new FakeShape("c", 0, 0, 1, 1, 1)]);
        ppt.DeleteShape("p", 1, "G");
        Assert.Equal(["c"], ppt.Slides[0].Select(s => s.Name));
    }
}
