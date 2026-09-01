using BckPPT.Rag;

namespace BckPPT.Tests;

public sealed class PptIndexerTests : IDisposable
{
    private readonly string _tempDirectory = Path.Combine(Path.GetTempPath(), "BckPPT.Tests", Guid.NewGuid().ToString("N"));

    [Fact]
    public async Task IndexFolder_IsIncremental_AndSlidesAreSearchable()
    {
        Directory.CreateDirectory(_tempDirectory);
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "icon.pptx"));
        Assert.True(File.Exists(fixture));
        File.Copy(fixture, Path.Combine(_tempDirectory, "fixture.pptx"));
        var repository = new TemplateRepository(Path.Combine(_tempDirectory, "index.db"));
        var indexer = new PptIndexer(repository);

        var first = await indexer.IndexFolderAsync(_tempDirectory);
        var second = await indexer.IndexFolderAsync(_tempDirectory);

        Assert.Equal(1, first.Files);
        Assert.True(first.Slides > 0);
        Assert.Equal(1, second.Skipped);
        Assert.Equal(first.Slides, repository.ReadAll().Count);
    }

    public void Dispose()
    {
        try { if (Directory.Exists(_tempDirectory)) Directory.Delete(_tempDirectory, true); } catch { }
    }
}
