using System.IO;

namespace BckPPT.Services;

/// <summary>
/// PowerPoint 연결 진단 로그(%LOCALAPPDATA%\BckSoft\BckPPT\logs\com-날짜.log). MCP 호출과 그 뒤 PowerPoint 상태,
/// 그동안 조용히 넘기던 COM 실패(문서 닫기 실패, PowerPoint 새로 띄우기)를 남긴다.
/// MCP로 고친 PPT를 닫고 탐색기에서 다시 열면 빈 PowerPoint만 뜨는 일이 있었는데 재현하지 못했다(2026-11-03).
/// 다시 생기면 그 시각 전후의 상태 줄로 어떤 문서가 창 없이 남았는지 본다.
/// </summary>
public static class PptDiagnosticLog
{
    private static readonly object Gate = new();

    public static string Directory { get; set; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BckSoft", "BckPPT", "logs");

    public static void Write(string kind, string message)
    {
        try
        {
            var now = DateTime.Now;
            lock (Gate)
            {
                System.IO.Directory.CreateDirectory(Directory);
                File.AppendAllText(Path.Combine(Directory, $"com-{now:yyyyMMdd}.log"),
                    $"{now:HH:mm:ss.fff} [{kind}] {message.ReplaceLineEndings(" ⏎ ")}{Environment.NewLine}");
            }
        }
        catch { }
    }
}
