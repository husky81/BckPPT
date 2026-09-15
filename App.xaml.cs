using System.IO;
using System.Windows;
using System.Windows.Threading;
using BckPPT.Application;
using BckPPT.Services;
using BckPPT.Infrastructure;

namespace BckPPT;

public partial class BckPptApp : System.Windows.Application
{
    private AppCompositionRoot? _compositionRoot;
    private SingleInstanceCoordinator? _singleInstance;
    private bool _staStarted;
    private static readonly string LogPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "BckSoft", "BckPPT", "crash.log");

    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);
        _singleInstance = new SingleInstanceCoordinator();
        if (!_singleInstance.IsPrimary)
        {
            var request = ReadExternalRequest(e.Args);
            if (request is not null)
            {
                try { SingleInstanceCoordinator.SendAsync(request).GetAwaiter().GetResult(); } catch { }
            }
            _singleInstance.Dispose();
            Environment.Exit(0);
        }
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

        _ = StaDispatcher.Instance;
        _staStarted = true;

        _compositionRoot = new AppCompositionRoot(Dispatcher, StartupOptions.Parse(e.Args));
        var chatWindow = _compositionRoot.CreateAgentChatWindow();
        MainWindow = chatWindow;
        MainWindow.Show();
        _compositionRoot.StartExternalEndpoint(chatWindow);

        if (ReadExternalRequest(e.Args) is { } startupRequest)
        {
            try
            {
                var request = System.Text.Json.JsonSerializer.Deserialize<ExternalCommandRequest>(startupRequest,
                    new System.Text.Json.JsonSerializerOptions(System.Text.Json.JsonSerializerDefaults.Web));
                if (request is not null) chatWindow.AcceptExternalRequest(request);
            }
            catch { }
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _compositionRoot?.Dispose();
        _singleInstance?.Dispose();
        if (_staStarted) StaDispatcher.Instance.Dispose();
        base.OnExit(e);
    }

    private static string? ReadExternalRequest(string[] args)
    {
        if (args.Length >= 2 && args[0] == "--request-json") return args[1];
        if (args.Length >= 2 && args[0] == "--request-file")
            return File.Exists(args[1]) ? File.ReadAllText(args[1]) : null;
        return null;
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
