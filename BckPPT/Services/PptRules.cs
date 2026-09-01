using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using CellBox2.BckPPT.Models;

namespace CellBox2.BckPPT.Services;

/// <summary>
/// AutoFit 파라미터, 허용 글꼴, 치환 규칙 등을 외부 JSON 파일로 관리.
/// %LOCALAPPDATA%\BckSoft\BckPPT\settings.json
/// </summary>
public class PptRules
{
    // ── 파일 경로 ───────────────────────────────────────────────

    public static readonly string ConfigPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "BckSoft", "BckPPT", "settings.json");

    // ── 현재 인스턴스 ────────────────────────────────────────────

    private static PptRules? _current;
    public static PptRules Current => _current ??= Load();

    // ── AutoFit 파라미터 ─────────────────────────────────────────

    public float MinFontSize        { get; set; } = 10f;
    public float OverflowTolerance  { get; set; } = 1f;
    public float SpacingMinimum     { get; set; } = -2f;
    public float SpaceAfterMinimum  { get; set; } = 0f;
    public float SpacingStep        { get; set; } = 0.5f;
    public float SpaceAfterStep     { get; set; } = 2f;

    // ── 글꼴 규칙 ────────────────────────────────────────────────

    public List<string> AllowedFonts { get; set; } =
    [
        "맑은 고딕", "Arial", "Calibri"
    ];

    public Dictionary<string, string> FontFallbackMap { get; set; } = new()
    {
        ["나눔고딕"]      = "맑은 고딕",
        ["나눔명조"]      = "맑은 고딕",
        ["함초롬바탕"]    = "맑은 고딕",
        ["굴림"]          = "맑은 고딕",
        ["Times New Roman"] = "Calibri",
    };

    public string DefaultFontFallback { get; set; } = "맑은 고딕";

    // ── 치환 규칙 ────────────────────────────────────────────────

    public List<ReplaceRule> ReplaceRules { get; set; } = [];

    // ── 파생 ─────────────────────────────────────────────────────

    [JsonIgnore]
    public HashSet<string> AllowedFontsSet
        => new(AllowedFonts, StringComparer.OrdinalIgnoreCase);

    [JsonIgnore]
    public Dictionary<string, string> FallbackMapCI
        => new(FontFallbackMap, StringComparer.OrdinalIgnoreCase);

    // ── Load / Save ──────────────────────────────────────────────

    public static PptRules Load()
    {
        try
        {
            if (File.Exists(ConfigPath))
            {
                var json  = File.ReadAllText(ConfigPath);
                var rules = JsonSerializer.Deserialize<PptRules>(json,
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (rules is not null)
                    return _current = rules;
            }
        }
        catch { }

        // 기본값으로 초기화 후 파일 생성
        var defaults = new PptRules();
        defaults.Save();
        return _current = defaults;
    }

    public void Save()
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
            var json = JsonSerializer.Serialize(this,
                new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(ConfigPath, json);
            _current = this;
        }
        catch { }
    }

    public static void Reload() => _current = null; // 다음 Current 접근 시 재로드
}
