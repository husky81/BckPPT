using System.IO;
using System.Text.Json;

namespace BckPPT.Application;

/// <summary>에이전트 콘솔의 도구 줄 한 칸: 한글 동작 이름과 인자 요약.</summary>
public sealed record ToolLine(string Label, string Detail, bool Hidden = false);

/// <summary>
/// 에이전트 콘솔에 보이는 글을 만든다. 도구 이름(remove_blank_lines)과 JSON 인자 대신 사람이 읽는 동작
/// ("빈 줄 정리 · 38–41쪽")을 보여 주고, 결과는 첫 줄만 다듬어 한 줄로 쓴다. 화면(WPF)과 떨어져 있어 테스트한다.
/// </summary>
public static class ConsoleFormatting
{
    public static ToolLine Describe(string tool, string argumentsJson)
    {
        JsonElement a = default;
        try { a = JsonDocument.Parse(string.IsNullOrWhiteSpace(argumentsJson) ? "{}" : argumentsJson).RootElement.Clone(); }
        catch (JsonException) { }

        string S(string name) => a.ValueKind == JsonValueKind.Object && a.TryGetProperty(name, out var v)
            ? v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : v.GetRawText() : "";
        int Count(string name) => a.ValueKind == JsonValueKind.Object && a.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Array
            ? v.GetArrayLength() : 0;
        string Page(string name) => S(name) is { Length: > 0 } p ? $"{p}쪽" : "";
        string Range(string from, string to) => S(to) is { Length: > 0 } t && t != S(from) ? $"{S(from)}–{t}쪽" : Page(from);
        string Shape() => S("shape") is { Length: > 0 } s ? $" {ShortShape(s)}" : "";
        static string Join(params string[] parts) => string.Join(" · ", parts.Where(p => p.Length > 0));

        return tool switch
        {
            "list_slides" => new("슬라이드 목록 보기", ""),
            "list_shapes" => new("도형 살펴보기", Page("slide")),
            "read_text" => new("글 읽기", Page("slide") + Shape()),
            "line_breaks" => new("줄 측정", Page("slide") + Shape()),
            "duplicate_slide" => new("쪽 복제", $"{S("source")}쪽 → {S("after")}쪽 뒤"),
            "new_blank_slide" => new("빈 쪽 만들기", Join($"{S("source")}쪽 양식", $"{S("after")}쪽 뒤")),
            "set_paragraphs" => new("문단 바꾸기", Page("slide") + Shape()),
            "set_texts" => new("상자 채우기", Join(Page("slide"), Count("items") > 0 ? $"{Count("items")}개" : "")),
            "replace_text" => new("글 바꾸기", Join(Page("slide"), $"'{Clip(S("find"), 16)}' → '{Clip(S("replace"), 16)}'")),
            "rename_shape" or "rename_shapes" => new("이름 붙이기", Page("slide")),
            "set_geometry" => new("위치·크기 맞추기", Page("slide") + Shape()),
            "delete_shape" => new("도형 지우기", Page("slide") + Shape()),
            "add_text_shape" => new("도형 추가", Join(Page("slide"), Clip(S("text"), 20))),
            "add_line" => new("선 긋기", Page("slide")),
            "add_table" => new("표 만들기", Join(Page("slide"), Count("rows") > 0 ? $"{Count("rows")}행" : "")),
            "group_shapes" => new("도형 묶기", Join(Page("slide"), Count("shapes") > 0 ? $"{Count("shapes")}개" : "")),
            "ungroup_shape" => new("그룹 풀기", Page("slide") + Shape()),
            "clear_slide" => new("쪽 비우기", Page("slide")),
            "draw_schedule" => new("일정표 그리기", Join(StyleName(S("style")), S("slide") is { Length: > 0 } ? Page("slide") : S("after") is { Length: > 0 } ? $"{S("after")}쪽 뒤" : "")),
            "remove_blank_lines" => new("빈 줄 정리", Range("from_slide", "to_slide")),
            "set_text_color" => new("글자색 바꾸기", Page("slide") + Shape()),
            "set_text_format" => new("글자 서식 바꾸기", Page("slide") + Shape()),
            "lint_slide" => new("레이아웃 검사", Page("slide")),
            "render_slide" => new("화면으로 확인", Page("slide")),
            "fill_cards" => new("카드 만들기", ""),
            "brief_source" => new("원문 요점 정리", S("document")),
            "source_outline" => new("원문 목차 보기", S("document")),
            "search_source" => new("원문 검색", $"'{Clip(S("query"), 24)}'"),
            "read_source" => new("원문 읽기", Join(S("document"), Range("from_page", "to_page"))),
            "finish" => new("마무리", "", Hidden: true),   // 요약은 답으로 따로 보인다
            _ => new(tool, ""),
        };
    }

    /// <summary>도구 결과의 첫 줄을 한 줄로 다듬는다. 오류는 "오류: " 머리를 뗀다.</summary>
    public static string ResultLine(string text, bool isError, int max = 90)
    {
        var line = text.Split('\n').Select(l => l.Trim()).FirstOrDefault(l => l.Length > 0) ?? (isError ? "실패" : "완료");
        if (line.StartsWith("오류:", StringComparison.Ordinal)) line = line[3..].Trim();
        // 목록 머리의 설명("38쪽 도형 64개 (주소 (이름) | left,top …)")은 괄호 앞까지만.
        if (line.Length > 40 && line.IndexOf(" (", StringComparison.Ordinal) is > 8 and var cut) line = line[..cut];
        return Clip(line, max);
    }

    /// <summary>"@123"이나 "그룹/이름" 주소를 짧게.</summary>
    public static string ShortShape(string address) =>
        address.StartsWith('@') ? "도형" : Clip(address.Contains('/') ? address[(address.LastIndexOf('/') + 1)..] : address, 18);

    private static string StyleName(string style) => style switch
    {
        "gantt" => "간트형", "inbar" or "in_bar" => "막대 안 이름형", "" or "table" => "표형", _ => style,
    };

    public static string Clip(string text, int max) => text.Length <= max ? text : text[..Math.Max(0, max - 1)].TrimEnd() + "…";

    /// <summary>
    /// 가운데를 줄여 너비에 맞춘다(앞뒤를 남긴다). 확장자는 끝에 남긴다. 예: "착수보고_…_042850.pptx".
    /// <paramref name="measure"/>는 글자열의 실제 표시 폭(px)이다.
    /// </summary>
    public static string MiddleEllipsis(string text, double maxWidth, Func<string, double> measure)
    {
        if (maxWidth <= 0) return "";
        if (measure(text) <= maxWidth) return text;
        string ext = Path.GetExtension(text);
        string stem = text[..^ext.Length];
        // 남길 글자 수를 이분 탐색한다. 앞은 조금 더 길게(파일명 앞부분이 더 알아보기 쉽다).
        int lo = 0, hi = stem.Length;
        string best = "…" + ext;
        while (lo <= hi)
        {
            int keep = (lo + hi) / 2;
            int head = (keep + 1) * 3 / 5, tail = keep - head;
            string candidate = stem[..head] + "…" + stem[^tail..] + ext;
            if (tail < 0 || head > stem.Length) { hi = keep - 1; continue; }
            if (measure(candidate) <= maxWidth) { best = candidate; lo = keep + 1; }
            else hi = keep - 1;
        }
        return best;
    }

    /// <summary>폴더 경로를 마지막 한두 단계만 남긴다. 예: "…\261006-기존 기획과제…\".</summary>
    public static string ShortenFolder(string? folder, int lastParts = 1, int partMax = 16)
    {
        if (string.IsNullOrEmpty(folder)) return "";
        var parts = folder.TrimEnd('\\', '/').Split('\\', '/').Where(p => p.Length > 0).ToList();
        var kept = parts.TakeLast(lastParts).Select(p => Clip(p, partMax));
        return (parts.Count > lastParts ? "…\\" : "") + string.Join("\\", kept) + "\\";
    }
}
