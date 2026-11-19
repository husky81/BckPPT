using System.Text;
using System.Text.Json.Nodes;
using BckPPT.Mcp;
using BckPPT.Services;

namespace BckPPT.Tests;

/// <summary>
/// 실제 PowerPoint로 export_pdf를 돌려 본다. BCKPPT_LIVE_PDF에 폴더를 주면 돈다:
/// test6.pptx(6쪽, 3쪽 숨김)를 사본으로 떠 내보내고 결과 글을 그 폴더의 pdf-live.log에 남긴다. 사용자 문서는 건드리지 않는다.
/// BCKPPT_LIVE_PDF_BIG에 큰 PPT 경로를 주면 그 파일도 읽기 전용으로 열어 걸린 시간을 잰다(PDF는 같은 폴더에 쓴다).
/// </summary>
[Collection("PptDiagnosticLog")]
public sealed class PdfExportLiveTests
{
    [Fact]
    public async Task ExportPdf_OnRealPowerPoint()
    {
        var dir = Environment.GetEnvironmentVariable("BCKPPT_LIVE_PDF");
        if (string.IsNullOrEmpty(dir)) return;
        PptDiagnosticLog.Directory = Path.Combine(dir, "logs");
        var deck = Path.Combine(dir, $"pdf-run-{DateTime.Now:HHmmss}.pptx");
        File.Copy(Path.Combine(dir, "test6.pptx"), deck, overwrite: true);
        var server = new McpServer(new McpToolCatalog(new PptComService()));
        var log = new StringBuilder();

        async Task<(bool Ok, string Text)> Call(JsonObject args)
        {
            var body = new JsonObject
            {
                ["jsonrpc"] = "2.0", ["id"] = 1, ["method"] = "tools/call",
                ["params"] = new JsonObject { ["name"] = "export_pdf", ["arguments"] = args },
            };
            log.AppendLine($"## export_pdf {args.ToJsonString()}");
            var reply = (await server.HandleBodyAsync(body.ToJsonString(), CancellationToken.None))!;
            var text = string.Join("\n", reply["result"]!["content"]!.AsArray().Select(c => c!["text"]?.GetValue<string>()));
            bool ok = !reply["result"]!["isError"]!.GetValue<bool>();
            log.AppendLine(text).AppendLine($"isError={!ok}");
            return (ok, text);
        }

        try
        {
            var full = await Call(new JsonObject { ["file"] = deck });
            Assert.True(full.Ok, full.Text);
            Assert.Contains("5쪽", full.Text);                       // 6쪽 중 숨긴 1쪽 뺌
            Assert.True(File.Exists(Path.ChangeExtension(deck, ".pdf")));

            var again = await Call(new JsonObject { ["file"] = deck });
            Assert.False(again.Ok);                                   // 같은 이름 있으면 멈춤
            Assert.Contains("overwrite", again.Text);

            var hidden = await Call(new JsonObject { ["file"] = deck, ["overwrite"] = true, ["include_hidden"] = true });
            Assert.True(hidden.Ok, hidden.Text);
            Assert.Contains("6쪽", hidden.Text);

            var one = await Call(new JsonObject { ["file"] = deck, ["slides"] = "2-4", ["output_path"] = Path.Combine(dir, "range.pdf"), ["overwrite"] = true });
            Assert.True(one.Ok, one.Text);
            Assert.Equal(2, CountPdfPages(Path.Combine(dir, "range.pdf")));

            var many = await Call(new JsonObject { ["file"] = deck, ["slides"] = new JsonArray(1, 2, 5, 6), ["output_path"] = Path.Combine(dir, "multi.pdf"), ["overwrite"] = true });
            Assert.True(many.Ok, many.Text);
            Assert.Equal(4, CountPdfPages(Path.Combine(dir, "multi.pdf")));

            Assert.Equal(6, CountPdfPages(Path.ChangeExtension(deck, ".pdf")));   // include_hidden으로 덮어쓴 판

            // 이미 열려 있고 저장 안 된 고침이 있는 문서: 그대로 내보내고 결과에 알린다. 문서는 그대로 남는다.
            var opened = Path.Combine(dir, $"pdf-open-{DateTime.Now:HHmmss}.pptx");
            File.Copy(Path.Combine(dir, "test6.pptx"), opened, overwrite: true);
            dynamic app = Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application")!)!;
            dynamic prs = app.Presentations.Open(opened, 0, 0, 0);
            try
            {
                prs.Slides[1].Shapes[1].TextFrame.TextRange.Text = "UNSAVED EDIT";
                var dirty = await Call(new JsonObject { ["file"] = opened, ["slides"] = "1" });
                Assert.True(dirty.Ok, dirty.Text);
                Assert.Contains("저장 안 된 고침", dirty.Text);
                Assert.Equal(1, CountPdfPages(Path.ChangeExtension(opened, ".pdf")));
                Assert.Equal(0, (int)prs.Saved);                         // 저장 표시를 건드리지 않았다
                Assert.Equal("UNSAVED EDIT", (string)prs.Slides[1].Shapes[1].TextFrame.TextRange.Text);
            }
            finally { prs.Saved = -1; prs.Close(); }

            var big = Environment.GetEnvironmentVariable("BCKPPT_LIVE_PDF_BIG");
            if (!string.IsNullOrEmpty(big))
            {
                var r = await Call(new JsonObject { ["file"] = big, ["output_path"] = Path.Combine(dir, "big.pdf"), ["overwrite"] = true });
                Assert.True(r.Ok, r.Text);
            }
        }
        finally
        {
            await File.WriteAllTextAsync(Path.Combine(dir, "pdf-live.log"), log.ToString());
        }
    }

    /// <summary>PDF의 쪽 객체(/Type /Page) 수. PowerPoint가 만든 PDF는 압축 안 된 사전에 적는다.</summary>
    private static int CountPdfPages(string path)
    {
        var text = Encoding.Latin1.GetString(File.ReadAllBytes(path));
        return System.Text.RegularExpressions.Regex.Matches(text, @"/Type\s*/Page(?![s\w])").Count;
    }
}
