using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace BckPPT.Services;

/// <summary>Read-only Restart Manager diagnostics. Never shuts down or restarts applications.</summary>
internal static class PptFileLockDiagnostics
{
    private const int MoreData = 234;
    private static readonly SemaphoreSlim Gate = new(1, 1);

    // Run outside the COM STA: a slow OS diagnostic must not delay document cleanup.
    public static void LogOwners(string path, string document)
    {
        if (!Gate.Wait(0))
        {
            PptDiagnosticLog.Event("file-lock-owners-skipped", new { path, document, reason = "probe-already-running" });
            return;
        }
        _ = Task.Run(() =>
        {
            var watch = Stopwatch.StartNew();
            try
            {
                PptDiagnosticLog.Event("file-lock-owners-start", new { path, document });
                var snapshot = ReadOwners(path);
                PptDiagnosticLog.Event("file-lock-owners", new { path, document, elapsedMs = watch.ElapsedMilliseconds, snapshot });
            }
            catch (Exception ex)
            {
                PptDiagnosticLog.Event("file-lock-owners-failed", new { path, document, elapsedMs = watch.ElapsedMilliseconds, error = ex.ToString() });
            }
            finally { Gate.Release(); }
        });
    }

    internal sealed record Owner(int Pid, string StartedAt, string AppName, string ServiceName,
        int ApplicationType, uint AppStatus, uint SessionId);
    internal sealed record Snapshot(string Stage, int Error, Owner[] Owners);

    // API references:
    // https://learn.microsoft.com/windows/win32/api/restartmanager/nf-restartmanager-rmgetlist
    // https://learn.microsoft.com/windows/win32/api/restartmanager/ns-restartmanager-rm_process_info
    internal static Snapshot ReadOwners(string path)
    {
        int result = RmStartSession(out uint session, 0, new StringBuilder(33));
        if (result != 0) return new("start", result, []);
        try
        {
            result = RmRegisterResources(session, 1, [path], 0, IntPtr.Zero, 0, IntPtr.Zero);
            if (result != 0) return new("register", result, []);
            uint count = 0;
            result = RmGetList(session, out uint needed, ref count, null, out _);
            for (int attempt = 0; result == MoreData && attempt < 3; attempt++)
            {
                if (needed > 4096) return new("list-limit", MoreData, []);
                var apps = new ProcessInfo[needed];
                count = (uint)apps.Length;
                result = RmGetList(session, out needed, ref count, apps, out _);
                if (result == 0)
                    return new("complete", 0, apps.Take((int)count).Select(app => new Owner(
                        app.Process.Pid, DateTime.FromFileTimeUtc(
                            ((long)(uint)app.Process.Started.dwHighDateTime << 32) | (uint)app.Process.Started.dwLowDateTime).ToString("O"),
                        app.AppName, app.ServiceName, app.ApplicationType, app.AppStatus, app.SessionId)).ToArray());
            }
            // An empty successful list does not prove the file is unlocked (cloud/driver locks may be absent).
            return new(result == 0 ? "complete" : "list", result, []);
        }
        finally
        {
            int ended = RmEndSession(session);
            if (ended != 0) PptDiagnosticLog.Event("file-lock-session-end-failed", new { path, error = ended });
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct UniqueProcess { public int Pid; public FILETIME Started; }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct ProcessInfo
    {
        public UniqueProcess Process;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string AppName;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string ServiceName;
        public int ApplicationType;
        public uint AppStatus;
        public uint SessionId;
        [MarshalAs(UnmanagedType.Bool)] public bool Restartable;
    }

    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmStartSession(out uint session, uint flags, StringBuilder key);
    [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)]
    private static extern int RmRegisterResources(uint session, uint files, string[] paths,
        uint applications, IntPtr processes, uint services, IntPtr serviceNames);
    [DllImport("rstrtmgr.dll")]
    private static extern int RmGetList(uint session, out uint needed, ref uint count,
        [In, Out] ProcessInfo[]? apps, out uint rebootReasons);
    [DllImport("rstrtmgr.dll")]
    private static extern int RmEndSession(uint session);
}
