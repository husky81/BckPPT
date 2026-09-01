using System.IO;
using System.Text.Json;
using BckPPT.Models;

namespace BckPPT.Services;

/// <summary>
/// Shape 역할 라벨을 JSON Lines로 저장.
/// %LOCALAPPDATA%\BckSoft\BckPPT\labels.jsonl
/// append-only, 동일 키 재라벨링 시 last-wins.
/// </summary>
public static class ShapeRoleLabelStore
{
    public static readonly string LabelPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "BckSoft", "BckPPT", "labels.jsonl");

    private static readonly JsonSerializerOptions _opts = new()
    {
        WriteIndented = false,
    };

    public static void Save(ShapeRoleLabel label)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LabelPath)!);
            var line = JsonSerializer.Serialize(label, _opts) + Environment.NewLine;
            File.AppendAllText(LabelPath, line);
        }
        catch { }
    }

    public static void SaveMany(IEnumerable<ShapeRoleLabel> labels)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LabelPath)!);
            var lines = labels.Select(l => JsonSerializer.Serialize(l, _opts));
            File.AppendAllLines(LabelPath, lines);
        }
        catch { }
    }

    /// <summary>모든 라벨 로드. 동일 키는 마지막 값 우선.</summary>
    public static Dictionary<(string Prs, int Slide, string Shape), ShapeRoleLabel> LoadAll()
    {
        var result = new Dictionary<(string Prs, int Slide, string Shape), ShapeRoleLabel>();

        try
        {
            if (!File.Exists(LabelPath)) return result;
            foreach (var line in File.ReadAllLines(LabelPath))
            {
                try
                {
                    var label = JsonSerializer.Deserialize<ShapeRoleLabel>(line, _opts);
                    if (label is null) continue;
                    result[(label.PresentationName, label.SlideIndex, label.ShapeName)] = label;
                }
                catch { }
            }
        }
        catch { }
        return result;
    }
}
