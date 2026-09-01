using System.IO;
using System.Text.Json;
using BckPPT.Models;

namespace BckPPT.Services;

/// <summary>
/// 변경 이력을 JSON Lines 형식으로 로컬 파일에 저장.
/// %LOCALAPPDATA%\BckSoft\BckPPT\history.jsonl
/// </summary>
public static class ChangeHistoryService
{
    public static readonly string HistoryPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "BckSoft", "BckPPT", "history.jsonl");

    private static readonly JsonSerializerOptions _opts = new()
    {
        WriteIndented = false,
    };

    // ── 기록 ─────────────────────────────────────────────────────

    public static void Append(ChangeHistoryEntry entry)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(HistoryPath)!);
            var line = JsonSerializer.Serialize(entry, _opts) + Environment.NewLine;
            File.AppendAllText(HistoryPath, line);
        }
        catch { }
    }

    public static void AppendMany(IEnumerable<ChangeHistoryEntry> entries)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(HistoryPath)!);
            var lines = entries.Select(e => JsonSerializer.Serialize(e, _opts));
            File.AppendAllLines(HistoryPath, lines);
        }
        catch { }
    }

    // ── 조회 ─────────────────────────────────────────────────────

    public static IReadOnlyList<ChangeHistoryEntry> LoadRecent(int count = 500)
    {
        try
        {
            if (!File.Exists(HistoryPath)) return [];
            return File.ReadAllLines(HistoryPath)
                .TakeLast(count)
                .Select(line =>
                {
                    try { return JsonSerializer.Deserialize<ChangeHistoryEntry>(line, _opts); }
                    catch { return null; }
                })
                .Where(e => e is not null)
                .Select(e => e!)
                .Reverse()
                .ToList();
        }
        catch { return []; }
    }

    // ── 파일 경로 열기 ───────────────────────────────────────────

    public static string GetHistoryDirectory()
        => Path.GetDirectoryName(HistoryPath)!;
}
