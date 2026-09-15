using System.Text.RegularExpressions;
using BckPPT.Models;

namespace BckPPT.Services;

public record ParsedChange(string Id, string NewText, bool Insert);

public record ParseResult(
    IReadOnlyList<ParsedChange> Changes,
    IReadOnlyList<string>       Warnings,
    IReadOnlyList<string>       UnmatchedLines);

/// <summary>
/// 외부 에이전트의 답변을 그대로 붙여넣어도 치환 대상만 추려내는 관용 파서.
/// LLM을 쓰지 않고 정규식만으로 결정적으로 동작한다.
///
/// 설계 전제: 에이전트에게 "설명하지 말라"고 묶는 대신, ID가 붙은 줄만 집어간다.
/// 따라서 산문·표·코드펜스가 섞여 있어도 안전하다.
/// </summary>
public static class PastedPlanParser
{
    /// <summary>S{슬라이드}-{도형}-{문단}, 뒤에 +면 문단 삽입.</summary>
    private static readonly Regex IdRegex = new(
        @"\bS(\d+)-(\d+)-(\d+)(\+)?",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    /// <summary>ID 뒤에 반드시 와야 하는 구분자. 없으면 산문 속 ID 언급으로 보고 버린다.</summary>
    private static readonly Regex SeparatorRegex = new(
        @"^(?:\s*[:：=]|\s*-{1,2}>|\s*→|\t|\x20{2,})\s*",
        RegexOptions.Compiled);

    private static readonly string[] EmptyMarkers =
        ["(빈 문단)", "(빈문단)", "(삭제)", "(비움)", "(empty)", "\"\"", "''"];

    public static ParseResult Parse(string pasted, IReadOnlyList<ParagraphEntry> known)
    {
        var byId      = known.ToDictionary(e => e.Id, StringComparer.OrdinalIgnoreCase);
        var picked    = new Dictionary<string, ParsedChange>(StringComparer.OrdinalIgnoreCase);
        var order     = new List<string>();
        var warnings  = new List<string>();
        var unmatched = new List<string>();

        if (string.IsNullOrWhiteSpace(pasted))
            return new ParseResult([], [], []);

        foreach (var raw in pasted.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n'))
        {
            var line = raw.TrimEnd();

            // 코드펜스는 경계로 쓰지 않고 그냥 버린다.
            // 에이전트마다 코드블록 안에 넣기도, 밖에 두기도 하기 때문.
            if (line.TrimStart().StartsWith("```", StringComparison.Ordinal)) continue;
            if (line.Length == 0) continue;

            var idMatch = IdRegex.Match(line);
            if (!idMatch.Success) continue;   // 설명 문장은 여기서 자연히 걸러진다

            bool insert = idMatch.Groups[4].Success;
            string id   = $"S{idMatch.Groups[1].Value}-{idMatch.Groups[2].Value}-{idMatch.Groups[3].Value}";

            string? value = IsTableRow(line)
                ? ExtractFromTableRow(line, idMatch, warnings, id)
                : ExtractFromPlainLine(line, idMatch);

            if (value is null) { unmatched.Add(line.Trim()); continue; }

            value = Normalize(value);

            if (!byId.ContainsKey(id))
            {
                warnings.Add($"{id} — 목록에 없는 ID (무시함)");
                continue;
            }

            // 같은 문단의 "치환"과 "뒤에 추가"는 서로 다른 작업이므로 따로 담는다.
            string key = insert ? id + "+" : id;

            // 에이전트가 "변경 전/후"를 나란히 쓰는 경우가 흔하므로 마지막 것을 채택한다.
            if (picked.ContainsKey(key))
                warnings.Add($"{key} — 여러 번 나와 마지막 값을 사용함");
            else
                order.Add(key);

            picked[key] = new ParsedChange(id, value, insert);
        }

        return new ParseResult(
            order.Select(id => picked[id]).ToList(),
            warnings,
            unmatched);
    }

    // ── 줄 형태별 값 추출 ────────────────────────────────────────

    private static bool IsTableRow(string line)
        => line.TrimStart().StartsWith('|');

    /// <summary>
    /// 마크다운 표 행에서 값을 뽑는다.
    /// 2열이면 나머지 칸, 그 이상이면 화살표 뒷부분 또는 마지막 칸을 쓴다.
    /// 다의적일 수 있으므로 경고를 남기고, 최종 판단은 미리보기 창에 맡긴다.
    /// </summary>
    private static string? ExtractFromTableRow(
        string line, Match idMatch, List<string> warnings, string id)
    {
        var cells = line.Trim().Trim('|').Split('|').Select(c => c.Trim()).ToList();
        int idCell = cells.FindIndex(c => IdRegex.IsMatch(c) && StripDecoration(c).Length <= 16);
        if (idCell < 0) return null;

        var rest = cells.Where((_, i) => i != idCell)
                        .Where(c => c.Length > 0)
                        .ToList();
        if (rest.Count == 0) return null;
        if (rest.Count == 1) return rest[0];

        warnings.Add($"{id} — 표에 칸이 여러 개라 마지막 칸을 새 텍스트로 해석함");
        var last = rest[^1];
        int arrow = last.LastIndexOf('→');
        return arrow >= 0 ? last[(arrow + 1)..] : last;
    }

    /// <summary>일반 줄. ID 뒤에 분명한 구분자가 있어야만 값으로 인정한다.</summary>
    private static string? ExtractFromPlainLine(string line, Match idMatch)
    {
        var tail = line[(idMatch.Index + idMatch.Length)..];

        // 감싼 백틱이 ID 바로 뒤에 오는 경우: `S1-12-3`: 텍스트
        if (tail.StartsWith('`')) tail = tail[1..];

        var sep = SeparatorRegex.Match(tail);
        if (!sep.Success) return null;   // "S1-12-3 는 유지" 같은 산문 언급을 차단

        var value = tail[sep.Length..];
        int arrow = value.LastIndexOf('→');
        if (arrow >= 0) value = value[(arrow + 1)..];
        return value;
    }

    // ── 값 정리 ──────────────────────────────────────────────────

    private static string Normalize(string value)
    {
        value = StripDecoration(value);

        if (EmptyMarkers.Any(m => value.Equals(m, StringComparison.OrdinalIgnoreCase)))
            return "";

        // 문단 안의 줄바꿈은 Shift+Enter(\v)로. \r은 문단을 쪼개 인덱스를 깨뜨리므로 금지.
        return value.Replace(@"\n", "\v").Replace('\r', ' ');
    }

    /// <summary>
    /// 마크다운 장식을 벗긴다. 굵게(**)를 불릿(*)으로 오인하지 않도록
    /// 감싼 표시를 먼저 풀고 나서 앞쪽 불릿을 제거한다.
    /// </summary>
    private static string StripDecoration(string s)
    {
        s = s.Trim();

        // 앞쪽 불릿·인용·헤딩. 굵게 표시의 **는 건드리지 않는다.
        s = Regex.Replace(s, @"^(?:[->#]+|\*(?!\*))+\s*", "").Trim();

        // 감싼 백틱과 굵게는 중첩될 수 있으므로 더 벗길 게 없을 때까지 반복.
        while (true)
        {
            if (s.Length >= 2 && s[0] == '`' && s[^1] == '`')
            {
                s = s[1..^1].Trim();
                continue;
            }
            if (s.Length >= 4 && s.StartsWith("**", StringComparison.Ordinal)
                              && s.EndsWith("**", StringComparison.Ordinal))
            {
                s = s[2..^2].Trim();
                continue;
            }
            break;
        }
        return s;
    }
}
