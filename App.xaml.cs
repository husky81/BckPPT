using System.IO;
using System.Windows;
using System.Windows.Threading;
using BckPPT.Application;
using BckPPT.Services;
using BckPPT.Infrastructure;
using BckPPT.Mcp;
using BckPPT.Views;

namespace BckPPT;

public partial class BckPptApp : System.Windows.Application
{
    private AppCompositionRoot? _compositionRoot;
    private SingleInstanceCoordinator? _singleInstance;
    private TrayIcon? _tray;
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
            // 이미 트레이에서 돌고 있다. 외부 작업이 없으면(다시 실행·CellBox 메뉴) 채팅창을 띄우게 한다.
            var request = ReadExternalRequest(e.Args) ?? System.Text.Json.JsonSerializer.Serialize(new
            {
                protocolVersion = "1.0", command = ExternalCommandValidator.ShowWindow, requestId = Guid.NewGuid().ToString(),
            });
            try { SingleInstanceCoordinator.SendAsync(request).GetAwaiter().GetResult(); } catch { }
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

        // 창 없이 트레이에서 시작한다. MCP 서버로 외부 에이전트를 받고, 채팅창은 필요할 때 띄운다.
        _compositionRoot = new AppCompositionRoot(Dispatcher, StartupOptions.Parse(e.Args));
        var mcpError = _compositionRoot.StartBackgroundServices(Dispatcher);
        var root = _compositionRoot;
        _tray = new TrayIcon(
            mcpError is null ? $"MCP: {McpServer.Endpoint}" : $"MCP를 켜지 못했습니다: {mcpError}",
            McpServer.ClaudeAddCommand,
            openChat: () => root.ShowChatWindow(),
            openEditor: root.OpenEditorWindow,
            exit: Shutdown);
        _tray.ShowBalloon("BckPPT가 트레이에서 실행 중입니다",
            (mcpError is null ? $"MCP {McpServer.Endpoint}" : "MCP를 켜지 못했습니다") +
            (root.IsLlmAvailable ? "" : " · LLM 키 없음(채팅 꺼짐)") + "\n두 번 누르면 채팅창이 열립니다.");

        if (ReadExternalRequest(e.Args) is { } startupRequest)
        {
            try
            {
                var request = System.Text.Json.JsonSerializer.Deserialize<ExternalCommandRequest>(startupRequest,
                    new System.Text.Json.JsonSerializerOptions(System.Text.Json.JsonSerializerDefaults.Web));
                if (request is not null) _compositionRoot.ShowChatWindow().AcceptExternalRequest(request);
            }
            catch { }
        }
    }

    protected override void OnExit(ExitEventArgs e)
    {
        _tray?.Dispose();
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
