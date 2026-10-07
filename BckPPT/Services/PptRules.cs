using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using BckPPT.Models;

namespace BckPPT.Services;

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

    // ── LayoutLint 기준 ──────────────────────────────────────────

    /// <summary>자동 줄바꿈된 문단의 마지막 줄이 이 글자 수 이하이면 고아 단어로 본다.</summary>
    public int   OrphanMaxChars          { get; set; } = 4;
    /// <summary>2줄 제목의 둘째 줄이 첫 줄 대비 이 비율 미만이면 균형을 맞추라고 한다.</summary>
    public float TitleSecondLineMinRatio { get; set; } = 0.3f;
    /// <summary>이름의 역할 부분(C2_title → title)에 이 문자열이 들어가면 제목으로 본다.</summary>
    public List<string> TitleRoles { get; set; } = ["title", "제목"];

    // ── 에이전트 ─────────────────────────────────────────────────

    /// <summary>카드 문안을 사용자 고급 모델에서 쓰게 한다(요청문 복사·답 붙여넣기).</summary>
    public bool ExternalCardWriting { get; set; }

    /// <summary>카드 길이 맞춤에 LLM(qwen)을 쓴다. 끄면 넘친 항목을 코드로만 뺀다.</summary>
    public bool LlmLengthFit { get; set; } = true;

    /// <summary>대상 PPT가 PowerPoint 창에 열려 있으면 그 창에서 바로 작업한다(작업 사본 없음, 저장은 사용자가).</summary>
    public bool WorkInOpenWindow { get; set; } = true;

    /// <summary>에이전트 콘솔 창 위치·크기("left,top,width,height"). 다음에 같은 자리에 연다.</summary>
    public string? ChatWindowBounds { get; set; }

    /// <summary>에이전트 콘솔을 항상 위에 둔다.</summary>
    public bool ChatTopmost { get; set; }

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
