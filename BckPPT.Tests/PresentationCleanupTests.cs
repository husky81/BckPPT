using System.Collections;
using System.Reflection;
using BckPPT.Services;

namespace BckPPT.Tests;

public sealed class PresentationCleanupTests : IDisposable
{
    private static readonly Type Service = typeof(PptComService);
    private static object Call(string name, params object[] args) => Service.GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, args)!;
    private static IList Owned => (IList)Service.GetField("Owned", BindingFlags.NonPublic | BindingFlags.Static)!.GetValue(null)!;
    private readonly string path = Path.Combine(Path.GetTempPath(), $"bckppt_로그_{Guid.NewGuid():N}.tmp");
    public PresentationCleanupTests() { File.WriteAllText(path, ""); }
    public void Dispose() { Owned.Clear(); File.Delete(path); }
    private FakePresentation Open(FakeApp app) => (FakePresentation)Call("OpenOwned", app, path, -1, 0, 0);

    [Fact]
    public void ExistingHiddenUserDocumentIsNeverClosed()
    {
        var app = new FakeApp();
        var user = app.Presentations.Open(path, 0, 0, 0);
        Assert.Same(user, Open(app));
        Call("SafeClose", user);
        Assert.Equal(0, user.CloseCalls);
        Assert.Empty(Owned);
    }

    [Fact]
    public void TransientFailureIsRetriedAndOwnershipRemovedAfterClose()
    {
        var app = new FakeApp();
        var prs = Open(app);
        prs.Failures = 1;
        Call("SafeClose", prs);
        Assert.Equal(2, prs.CloseCalls);
        Assert.Empty(Owned);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void VisibleOrUnsavedDocumentsArePreserved(bool visible)
    {
        var prs = Open(new FakeApp());
        if (visible) prs.Windows.Count = 1; else prs.Saved = 0;
        Call("SafeClose", prs);
        Assert.Equal(0, prs.CloseCalls);
        Assert.Equal(0, prs.SaveCalls);
    }

    [Fact]
    public void FailedCloseRetainsOwnershipForLaterRetry()
    {
        var prs = Open(new FakeApp());
        prs.Failures = 3;
        Call("SafeClose", prs);
        Assert.Single(Owned);
        Call("RetryPendingCloses");
        Assert.Empty(Owned);
        Assert.Equal(4, prs.CloseCalls);
    }

    [Fact]
    public void SaveFailurePreservesDirtyDocumentAndCanRecover()
    {
        var prs = Open(new FakeApp());
        Owned[0]!.GetType().GetField("DeferredSave")!.SetValue(Owned[0], true);
        prs.Saved = 0;
        prs.SaveFails = true;
        Call("SafeClose", prs);
        Assert.Equal(0, prs.CloseCalls);
        Assert.Single(Owned);
        prs.SaveFails = false;
        Call("RetryPendingCloses");
        Assert.Empty(Owned);
        Assert.Equal(1, prs.CloseCalls);
    }

    [Fact]
    public void NestedHandlesCloseOnlyAfterLastDisposeAndDisposeIsIdempotent()
    {
        var prs = Open(new FakeApp());
        var owned = Owned[0]!;
        owned.GetType().GetField("Leases")!.SetValue(owned, 2);
        IDisposable Handle()
        {
            var type = Service.GetNestedType("OpenHandle", BindingFlags.NonPublic)!;
            var handle = (IDisposable)Activator.CreateInstance(type, nonPublic: true)!;
            type.GetField("Owned")!.SetValue(handle, owned);
            return handle;
        }
        var first = Handle();
        var second = Handle();
        first.Dispose();
        first.Dispose();
        Assert.Equal(0, prs.CloseCalls);
        second.Dispose();
        Assert.Equal(1, prs.CloseCalls);
        Assert.Empty(Owned);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExceptionAndCancellationStillCloseOwnedDocument(bool cancelled)
    {
        var prs = Open(new FakeApp());
        prs.Failures = 3;
        Call("SafeClose", prs);
        Assert.Single(Owned);
        var run = Service.GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .Single(m => m.Name == "RunOnSta" && !m.IsGenericMethod);
        Action operation = () =>
        {
            if (cancelled) throw new OperationCanceledException();
            throw new InvalidOperationException("operation failed");
        };
        Assert.Throws<TargetInvocationException>(() => run.Invoke(null, new object[] { operation, "test" }));
        Assert.Equal(4, prs.CloseCalls);
        Assert.Empty(Owned);
    }

    [Fact]
    public void LifecycleLogsCorrelateDocumentAndRecordSaveFailureStage()
    {
        var prs = Open(new FakeApp());
        var owned = Owned[0]!;
        string id = (string)owned.GetType().GetField("Id")!.GetValue(owned)!;
        owned.GetType().GetField("DeferredSave")!.SetValue(owned, true);
        prs.Saved = 0;
        prs.SaveFails = true;
        Call("SafeClose", prs);
        prs.SaveFails = false;
        Call("RetryPendingCloses");
        string logPath = Path.Combine(PptDiagnosticLog.Directory, $"com-{DateTime.Now:yyyyMMdd}.log");
        using var stream = new FileStream(logPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var reader = new StreamReader(stream);
        var lines = reader.ReadToEnd().Split('\n').Where(line => line.Contains(id)).ToArray();
        Assert.Contains(lines, line => line.Contains("[opened-owned]"));
        Assert.Contains(lines, line => line.Contains("[close-failed]") && line.Contains("stage=save"));
        Assert.Contains(lines, line => line.Contains("[close-pending-retry]"));
        Assert.Contains(lines, line => line.Contains("[close-complete]"));
        Assert.Contains(lines, line => line.Contains("[file-release-verified]"));
        Assert.All(lines, line => Assert.Contains($"host={Environment.ProcessId}", line));
        Assert.Contains(lines, line => line.Contains(Path.GetFileName(path))); // Korean filenames remain searchable.
    }

    [Fact]
    public void DiagnosticLinesIncludeRunIdentityAndFlattenMultilineErrors()
    {
        var method = typeof(PptDiagnosticLog).GetMethod("FormatLine", BindingFlags.NonPublic | BindingFlags.Static)!;
        string line = (string)method.Invoke(null, new object[] { "test", "first\r\nsecond", DateTimeOffset.Parse("2026-11-09T15:00:00+09:00") })!;
        Assert.StartsWith("2026-11-09T15:00:00.000+09:00 [test]", line);
        Assert.Contains("run=", line);
        Assert.Contains("seq=", line);
        Assert.Contains("thread=", line);
        Assert.Contains("first ⏎ second", line);
        Assert.DoesNotContain("\n", line.TrimEnd('\r', '\n'));
    }

    [Fact]
    public void RestartManagerIdentifiesOurTemporaryFileHandleWithoutClosingIt()
    {
        using var held = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite);
        var type = Service.Assembly.GetType("BckPPT.Services.PptFileLockDiagnostics")!;
        var snapshot = type.GetMethod("ReadOwners", BindingFlags.NonPublic | BindingFlags.Static)!.Invoke(null, new object[] { path })!;
        Assert.Equal(0, (int)snapshot.GetType().GetProperty("Error")!.GetValue(snapshot)!);
        var owners = (IEnumerable)snapshot.GetType().GetProperty("Owners")!.GetValue(snapshot)!;
        Assert.Contains(owners.Cast<object>(), owner => (int)owner.GetType().GetProperty("Pid")!.GetValue(owner)! == Environment.ProcessId);
        Assert.True(held.CanRead);
    }

    public sealed class FakeApp
    {
        public FakeCollection Presentations { get; }
        public FakeApp() { Presentations = new(this); }
    }
    public sealed class FakeCollection(FakeApp app) : List<FakePresentation>
    {
        public FakePresentation Open(string path, int ro, int untitled, int window)
        {
            var prs = new FakePresentation { FullName = path, Application = app };
            Add(prs);
            return prs;
        }
    }
    public sealed class FakeWindows { public int Count { get; set; } }
    public sealed class FakePresentation
    {
        public string FullName { get; set; } = "";
        public FakeApp Application { get; set; } = null!;
        public FakeWindows Windows { get; } = new();
        public int Saved { get; set; } = -1;
        public int CloseCalls, SaveCalls, Failures;
        public bool SaveFails;
        public void Save() { SaveCalls++; if (SaveFails) throw new IOException("save failure"); Saved = -1; }
        public void Close() { CloseCalls++; if (Failures-- > 0) throw new IOException("busy"); Application.Presentations.Remove(this); }
    }
}
