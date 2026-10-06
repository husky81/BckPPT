using System.IO;
using BckPPT.Agent.Cards;
using BckPPT.Application;

namespace BckPPT.Tests;

public sealed class AgentWorkflowTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "bckppt-wf-tests-" + Guid.NewGuid().ToString("N")[..8], "기존 기획보고서");

    public AgentWorkflowTests()
    {
        Directory.CreateDirectory(_dir);
        foreach (var name in new[] { "04. 노후 도로시설물 기획.md", "04. 노후 도로시설물 기획.pdf", "건설로보틱스.pdf", "목록.xlsx" })
            File.WriteAllText(Path.Combine(_dir, name), "x");
    }

    [Fact]
    public void 요청에_적힌_폴더_경로를_찾는다()
    {
        // 2026-10-07 채팅: 따옴표로 감싼 폴더 경로를 적었는데 첨부되지 않았다.
        var quoted = $"\"{_dir}\" 이 폴더에 있는 md 파일 하나 마다 카드를 만들어줘";
        Assert.Equal([_dir], AgentWorkflow.MentionedSourcePaths(quoted));

        var bare = $"{_dir} 이 폴더로 카드를 만들어줘";      // 따옴표 없이 공백 있는 경로
        Assert.Equal([_dir], AgentWorkflow.MentionedSourcePaths(bare));

        Assert.Empty(AgentWorkflow.MentionedSourcePaths(@"C:\없는\경로 로 해줘"));
    }

    [Fact]
    public void 같은_이름이면_md를_pdf보다_고른다()
    {
        var picked = AgentWorkflow.PreferTextVersions(Directory.EnumerateFiles(_dir)
            .Where(f => BckPPT.Rag.SourceLibrary.SupportedExtensions.Contains(Path.GetExtension(f))));
        Assert.Equal(["04. 노후 도로시설물 기획.md", "건설로보틱스.pdf"], picked.Select(Path.GetFileName));
    }

    [Fact]
    public void 한_쪽에_넘치는_카드는_템플릿_카드_수대로_나눈다()
    {
        var cards = Enumerable.Range(1, 13).Select(i => new CardSpec($"{i}", $"과제{i}")).ToList();
        var job = new CardJob(35, 37, [new CardSlideSpec(cards, "4. 기존 기획과제 연계(1/2)")]);

        var split = CardJobExtractor.SplitToCapacity(job, 4);

        Assert.Equal([4, 4, 4, 1], split.Slides.Select(s => s.Cards.Count));
        Assert.Equal(["4. 기존 기획과제 연계(1/4)", "4. 기존 기획과제 연계(2/4)", "4. 기존 기획과제 연계(3/4)", "4. 기존 기획과제 연계(4/4)"],
            split.Slides.Select(s => s.Title));
    }

    [Fact]
    public void 쪽_구분_없이_cards만_준_사양도_읽는다()
    {
        var job = CardJobExtractor.Parse("""{"template_slide": 35, "after": 37, "cards": [{"number": "1", "title": "가"}, {"number": "2", "title": "나"}]}""");
        Assert.Equal(2, Assert.Single(job.Slides).Cards.Count);
    }

    public void Dispose()
    {
        try { Directory.Delete(Path.GetDirectoryName(_dir)!, true); } catch { }
    }
}
