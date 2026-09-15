using System.IO;
using BckPPT.Application;

namespace BckPPT.Tests;

public class StartupOptionsTests
{
    [Fact]
    public void Parse_NoArguments_ReturnsEmpty()
    {
        Assert.Null(StartupOptions.Parse([]).SyncFolder);
        Assert.Null(StartupOptions.Parse(null).SyncFolder);
    }

    [Fact]
    public void Parse_SeparateValue_ReadsPath()
    {
        var options = StartupOptions.Parse(["--sync-folder", @"C:\CellBox\Sync"]);
        Assert.Equal(@"C:\CellBox\Sync", options.SyncFolder);
    }

    [Fact]
    public void Parse_EqualsForm_ReadsPath()
    {
        var options = StartupOptions.Parse([@"--sync-folder=C:\CellBox\Sync"]);
        Assert.Equal(@"C:\CellBox\Sync", options.SyncFolder);
    }

    [Fact]
    public void Parse_IgnoresUnrelatedArguments()
    {
        var options = StartupOptions.Parse(["--request-json", "{}", "--sync-folder", @"C:\Sync"]);
        Assert.Equal(@"C:\Sync", options.SyncFolder);
    }

    [Fact]
    public void Parse_FlagWithoutValue_ReturnsEmpty()
    {
        Assert.Null(StartupOptions.Parse(["--sync-folder"]).SyncFolder);
        Assert.Null(StartupOptions.Parse(["--sync-folder", "   "]).SyncFolder);
    }

    [Fact]
    public void ResolveInitialDirectory_MissingFolder_ReturnsNull()
    {
        var options = new StartupOptions(Path.Combine(Path.GetTempPath(), "bckppt-없는폴더-" + Guid.NewGuid()));
        Assert.Null(options.ResolveInitialDirectory());
    }

    [Fact]
    public void ResolveInitialDirectory_ExistingFolder_ReturnsFullPath()
    {
        var folder = Directory.CreateTempSubdirectory("bckppt-sync-");
        try
        {
            var options = new StartupOptions(folder.FullName);
            Assert.Equal(folder.FullName, options.ResolveInitialDirectory());
        }
        finally { folder.Delete(recursive: true); }
    }
}
