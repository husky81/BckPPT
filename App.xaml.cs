using System.IO;
using System.Windows;
using System.Windows.Threading;
using CellBox2.BckPPT.Services;

namespace CellBox2;

public partial class BckPptApp : Application
{
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "BckSoft", "BckPPT", "crash.log");

    protected override void OnStartup(StartupEventArgs e)
    {
        // 미처리 예외 → 파일 기록 + 다이얼로그
        DispatcherUnhandledException += (_, ex) =>
        {
            LogCrash("DispatcherUnhandledException", ex.Exception);
            ex.Handled = true;
            MessageBox.Show($"오류: {ex.Exception.Message}\n\n{ex.Exception.StackTrace}",
                "BckPPT 오류", MessageBoxButton.OK, MessageBoxImage.Error);
        };
        AppDomain.CurrentDomain.UnhandledException += (_, ex) =>
            LogCrash("UnhandledException", ex.ExceptionObject as Exception);

        base.OnStartup(e);
        _ = StaDispatcher.Instance;
    }

    protected override void OnExit(ExitEventArgs e)
    {
        StaDispatcher.Instance.Dispose();
        base.OnExit(e);
    }

    private static void LogCrash(string source, Exception? ex)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogPath)!);
            File.AppendAllText(LogPath,
                $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {source}\n" +
                $"{ex?.GetType().FullName}: {ex?.Message}\n" +
                $"{ex?.StackTrace}\n\n");
        }
        catch { }
    }
}
