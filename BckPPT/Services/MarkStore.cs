using System.IO;
using System.Text.Json;
using BckPPT.Models;

namespace BckPPT.Services;

/// <summary>
/// 빨간색 표시의 원래 글자색 스냅샷을 파일에 보관한다.
/// BckPPT를 껐다 켜도 "표시 해제"가 동작하게 하려면 메모리에만 둘 수 없다.
/// %LOCALAPPDATA%\BckSoft\BckPPT\marks.json
/// </summary>
public static class MarkStore
{
    private static readonly string StorePath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "BckSoft", "BckPPT", "marks.json");

    private static Dictionary<string, List<ParagraphColorSnapshot>> Load()
    {
        try
        {
            if (!File.Exists(StorePath)) return [];
            return JsonSerializer.Deserialize<Dictionary<string, List<ParagraphColorSnapshot>>>(
                File.ReadAllText(StorePath)) ?? [];
        }
        catch { return []; }
    }

    private static void Save(Dictionary<string, List<ParagraphColorSnapshot>> map)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(StorePath)!);
            File.WriteAllText(StorePath, JsonSerializer.Serialize(map));
        }
        catch { }
    }

    /// <summary>같은 문단을 다시 표시해도 최초의 원래 색을 잃지 않도록 기존 항목을 우선한다.</summary>
    public static void Add(string presentationPath, IReadOnlyList<ParagraphColorSnapshot> snapshots)
    {
        if (snapshots.Count == 0) return;
        var map  = Load();
        var list = map.TryGetValue(presentationPath, out var existing) ? existing : [];

        var seen = list.Select(Key).ToHashSet();
        list.AddRange(snapshots.Where(s => !seen.Contains(Key(s))));

        map[presentationPath] = list;
        Save(map);
    }

    public static IReadOnlyList<ParagraphColorSnapshot> Get(string presentationPath)
        => Load().TryGetValue(presentationPath, out var list) ? list : [];

    public static void Clear(string presentationPath)
    {
        var map = Load();
        if (map.Remove(presentationPath)) Save(map);
    }

    private static string Key(ParagraphColorSnapshot s)
        => $"{s.SlideIndex}|{s.ShapeName}|{s.ParagraphIndex}";
}
