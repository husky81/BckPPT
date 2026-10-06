using System.IO;
using System.Text;
using BckPPT.Agent.Tools;
using BckPPT.Rag;
using Xunit.Abstractions;

namespace BckPPT.Tests;

public sealed class SourceLibraryTests(ITestOutputHelper output) : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "bckppt-source-tests-" + Guid.NewGuid().ToString("N")[..8]);

    private SourceLibrary Library() => new(Path.Combine(_dir, "sources.db"));

    private string Write(string name, string content)
    {
        Directory.CreateDirectory(_dir);
        var path = Path.Combine(_dir, name);
        File.WriteAllText(path, content, Encoding.UTF8);
        return path;
    }

    /// <summary>400개 구간짜리 보고서. 대부분은 잡음이고 "연구개발 내용" 절은 한 곳뿐이다.</summary>
    private string LargeReport()
    {
        var sb = new StringBuilder("# 기획보고서\n\n");
        for (int i = 1; i <= 400; i++)
        {
            if (i == 237)
            {
                sb.Append("## 3. 연구개발 내용\n\n");
                sb.Append("본 과제의 연구개발 내용은 배터리 열폭주 조기 감지 센서와 예측 알고리즘의 개발이다.\n\n");
            }
            else
            {
                sb.Append($"## 부록 {i}\n\n");
                sb.Append($"사업 추진 일정과 예산 집행 계획 {i}. 연구 인력 운영 방안과 기관 협력 체계를 정리한다. ");
                sb.Append(i % 7 == 0 ? "연구 성과 확산 계획을 포함한다.\n\n" : "행정 사항을 기술한다.\n\n");
            }
        }
        return Write("기획보고서_A.md", sb.ToString());
    }

    [Fact]
    public void 큰_보고서에서_연구개발_내용_절을_상위에서_찾는다()
    {
        var library = Library();
        Assert.Equal(401, library.AddFile(LargeReport()));   // 맨 앞 제목 구간 + 400

        var hits = library.Search("연구개발 내용", 5);

        Assert.NotEmpty(hits);
        Assert.Equal("기획보고서_A.md", hits[0].Document);
        Assert.Contains("〈3. 연구개발 내용〉", hits[0].Snippet);
        Assert.Contains("열폭주", hits[0].Snippet);
    }

    [Fact]
    public void 조사가_붙은_한글도_찾는다()
    {
        var library = Library();
        library.AddFile(Write("rfp.md", "# 과업\n\n본 사업의 최종목표는 플랫폼 구축이다.\n\n# 기타\n\n관련 없음\n"));
        var hit = Assert.Single(library.Search("최종목표", 5));
        Assert.Equal(1, hit.Page);
    }

    [Fact]
    public void 두_글자_낱말도_찾는다()
    {
        // trigram은 3글자 미만을 못 찾으므로 LIKE로 보완한다.
        var library = Library();
        library.AddFile(Write("a.md", "# 개요\n\n짧은 절\n\n# 예산\n\n돈 이야기\n"));
        var hit = Assert.Single(library.Search("개요", 5));
        Assert.Contains("〈개요〉", hit.Snippet);
    }

    [Fact]
    public void 목차는_절_제목과_쪽_범위를_보여준다()
    {
        var library = Library();
        var body = string.Join("\n\n", Enumerable.Range(1, 30).Select(i => new string('나', 150) + i));
        library.AddFile(Write("e.md", "# 요약문\n\n짧음\n\n# 연구개발 내용\n\n" + body + "\n\n# 부록\n\n끝\n"));

        var outline = library.Outline("e.md");

        Assert.Contains("p.1 요약문", outline);
        Assert.Matches(@"p\.2–\d+ 연구개발 내용", outline);
        Assert.Contains("부록", outline);
        Assert.Contains("e.md (", library.Outline(null));
    }

    [Fact]
    public void 지정한_쪽을_읽는다()
    {
        var library = Library();
        library.AddFile(Write("b.md", "# 하나\n\n첫째 본문\n\n# 둘\n\n둘째 본문\n\n# 셋\n\n셋째 본문\n"));
        var text = library.Read("b", 2, 3);
        Assert.Contains("b.md p.2–3 (전체 3쪽)", text);
        Assert.Contains("둘째 본문", text);
        Assert.Contains("셋째 본문", text);
        Assert.DoesNotContain("첫째 본문", text);
    }

    [Fact]
    public void 없는_문서는_있는_목록과_함께_알린다()
    {
        var library = Library();
        library.AddFile(Write("c.md", "본문"));
        var ex = Assert.Throws<AgentToolException>(() => library.Read("없는문서", 1, 1));
        Assert.Contains("c.md", ex.Message);
    }

    [Fact]
    public void 같은_파일은_다시_색인하지_않고_바뀌면_교체한다()
    {
        var library = Library();
        var path = Write("d.md", "# 하나\n\n옛 내용\n");
        library.AddFile(path);
        library.AddFile(path);
        Assert.Single(library.Documents());

        File.WriteAllText(path, "# 하나\n\n새 내용 열폭주\n", Encoding.UTF8);
        library.AddFile(path);
        Assert.Single(library.Search("열폭주", 5));
        Assert.DoesNotContain(library.Search("옛 내용", 5), h => h.Snippet.Contains("옛"));
    }

    [Fact]
    public void 범위를_정하면_첨부한_문서만_찾는다()
    {
        var library = Library();
        var a = Write("a.md", "# 하나\n\n열폭주 감지\n");
        var b = Write("b.md", "# 둘\n\n열폭주 예측\n");
        library.AddFile(a);
        library.AddFile(b);

        var scoped = library.WithScope([b]);

        Assert.Equal("b.md", Assert.Single(scoped.Search("열폭주", 5)).Document);
        Assert.Single(scoped.Documents());
        Assert.Throws<AgentToolException>(() => scoped.Read("a.md", 1, 1));
        Assert.Equal(2, library.Search("열폭주", 5).Count);
    }

    [Fact]
    public void 긴_절은_빈_줄에서_나눈다()
    {
        var body = string.Join("\n\n", Enumerable.Range(1, 30).Select(i => new string('가', 150) + i));
        var pages = SourceLibrary.SplitText("# 긴 절\n\n" + body);
        Assert.True(pages.Count > 1);
        Assert.All(pages, p => Assert.True(p.Text.Length <= SourceLibrary.MaxSectionChars));
        Assert.All(pages, p => Assert.Equal("긴 절", p.Heading));
    }

    [Fact]
    public void 폼피드가_있으면_쪽_경계로_쓴다()
    {
        var pages = SourceLibrary.SplitText("1쪽 내용\f2쪽 내용\f3쪽 내용");
        Assert.Equal(3, pages.Count);
        Assert.Equal("2쪽 내용", pages[1].Text);
    }

    /// <summary>실제 원문으로 확인. BCKPPT_LIVE_SOURCE(파일)와 BCKPPT_LIVE_QUERY가 있을 때만 실행한다.</summary>
    [Fact]
    public void 실문서_검색()
    {
        var file = Environment.GetEnvironmentVariable("BCKPPT_LIVE_SOURCE");
        if (string.IsNullOrEmpty(file)) return;
        var library = Library();
        var started = DateTime.Now;
        int pages = library.AddFile(file);
        output.WriteLine($"{Path.GetFileName(file)}: {pages}쪽 색인 {(DateTime.Now - started).TotalSeconds:0.0}초");
        foreach (var query in (Environment.GetEnvironmentVariable("BCKPPT_LIVE_QUERY") ?? "과업 내용").Split('|'))
        {
            output.WriteLine($"\n## {query}");
            foreach (var hit in library.Search(query, 5))
                output.WriteLine($"p.{hit.Page} {hit.Snippet[..Math.Min(160, hit.Snippet.Length)]}");
        }
    }

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); } catch { }
    }
}
