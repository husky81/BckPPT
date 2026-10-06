using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using BckPPT.Models;
using BckPPT.PowerPoint;

namespace BckPPT.Agent.Cards;

/// <summary>카드 한 장의 사양. Document는 원문 별칭(D1 등)이다.</summary>
public sealed record CardSpec(string Number, string Title, string? Tag = null, string? Document = null);

/// <summary>만들 슬라이드 하나. Title을 주면 슬라이드 제목을 바꾼다.</summary>
public sealed record CardSlideSpec(IReadOnlyList<CardSpec> Cards, string? Title = null);

/// <summary>
/// 템플릿 카드 슬라이드를 복제해 카드를 채우는 작업. 배치·이름·삭제·폭은 코드가 하고,
/// LLM은 원문 요점 정리(문서마다 1회), 문안 작성(전체 1회), 길이 맞춤(넘친 상자만)에만 쓴다.
/// </summary>
public sealed record CardJob(
    int TemplateSlide,
    int InsertAfter,
    IReadOnlyList<CardSlideSpec> Slides,
    string? Header = null,
    string? Legend = null);

/// <summary>상자 하나의 글자 예산. 템플릿 상자의 실제 줄 수와 한 줄 글자 수로 정한다.</summary>
public sealed record TextBudget(string Role, string Label, int Lines, int CharsPerLine, int Paragraphs)
{
    public int TotalChars => Lines * CharsPerLine;

    /// <summary>PowerPoint가 나눈 줄로 예산을 잰다. 자동 줄바꿈된 줄이 꽉 찬 줄이므로 그 길이를 한 줄 글자 수로 본다.</summary>
    public static TextBudget Measure(string role, string label, IReadOnlyList<TextLineInfo> lines)
    {
        int Len(TextLineInfo l) => l.Text.Trim().Length;
        var full = lines.Where(l => l.BreakType == LineBreakType.AutoWrap).ToList();
        int perLine = full.Count > 0 ? full.Max(Len) : lines.Count > 0 ? Math.Max(lines.Max(Len), 10) : 20;
        int paragraphs = Math.Max(1, lines.Count(l => l.BreakType is LineBreakType.Enter or LineBreakType.None));
        return new TextBudget(role, label, Math.Max(1, lines.Count), perLine, paragraphs);
    }

    public string Describe() =>
        $"{Label}: 전체 {Lines}줄 이내(한 줄 약 {CharsPerLine}자, 템플릿은 항목 {Paragraphs}개)";
}

/// <summary>LLM이 쓴 카드 문안. Sections는 역할 키(overview 등)별 항목이다.</summary>
public sealed record CardDraft(string Number, string Title, IReadOnlyDictionary<string, IReadOnlyList<string>> Sections);

public static class CardFillPrompts
{
    public const string WriterSystem = """
        너는 국가 R&D 착수보고 슬라이드의 카드 문안을 쓰는 편집자다.
        각 과제의 원문 요점을 받아 카드에 들어갈 문안을 쓴다.

        규칙
        - 원문 요점에 있는 사실만 쓴다. 숫자·고유명사를 바꾸지 않는다.
        - 개조식 명사형으로 끝낸다("~개발", "~구축", "~부재", "~곤란"). 문장 끝에 마침표를 쓰지 않는다.
        - 상자마다 주어진 줄 예산을 지킨다. 항목 하나의 길이가 한 줄 글자 수를 넘으면 줄을 하나 더 쓰는 것으로 계산한다.
        - 한 단어만 다음 줄로 넘어가지 않게, 항목 길이를 한 줄 글자 수의 배수에 가깝게 맞추거나 한 줄 안에 끝낸다.
        - 제목은 과제명을 그대로 쓰되 "기획", "기획보고서" 같은 문서 종류는 뺀다.
        - 답은 JSON 하나만 쓴다. 설명·머리말·코드블록 표시를 붙이지 않는다.
        """;

    /// <summary>문안 작성 요청. 카드마다 원문 요점과, 템플릿에서 잰 상자별 예산을 준다.</summary>
    public static string WriterRequest(
        IReadOnlyList<(CardSpec Card, string Brief)> cards, TextBudget title, IReadOnlyList<TextBudget> sections,
        CardLayoutSample? sample)
    {
        var sb = new StringBuilder();
        sb.Append("상자별 줄 예산\n");
        sb.Append($"- 제목(title): {title.Lines}줄 이내, 한 줄 약 {title.CharsPerLine}자\n");
        foreach (var s in sections) sb.Append($"- {s.Describe()} → 키 \"{s.Role}\"\n");
        if (sample is not null)
        {
            sb.Append("\n템플릿 카드 예시(문체·길이 참고, 내용은 쓰지 않는다)\n");
            sb.Append($"- 제목: {sample.Title}\n");
            foreach (var (label, items) in sample.Sections) sb.Append($"- {label}: {string.Join(" / ", items)}\n");
        }
        sb.Append("\n출력 형식\n{\"cards\": [{\"number\": \"1\", \"title\": \"…\"");
        foreach (var s in sections) sb.Append($", \"{s.Role}\": [\"항목\", …]");
        sb.Append("}, …]}\n\n카드\n");
        foreach (var (card, brief) in cards)
        {
            sb.Append($"\n### 카드 번호 {card.Number} — {card.Title}\n");
            sb.Append(string.IsNullOrWhiteSpace(brief) ? "(원문 요점 없음: 과제명만으로 일반적인 내용을 지어내지 말고 항목을 비워 둔다)\n" : brief.Trim() + "\n");
        }
        return sb.ToString();
    }

    public const string FitSystem = """
        너는 슬라이드 상자의 문구 길이를 맞추는 교정자다. 뜻은 유지하고 길이만 맞춘다.
        개조식 명사형을 유지한다. 답은 JSON 배열(문자열 항목들) 하나만 쓴다.
        """;

    /// <summary>넘친 상자 하나를 고치는 요청. 실제 측정 결과를 그대로 보여준다.</summary>
    public static string FitRequest(TextBudget budget, IReadOnlyList<string> items, IReadOnlyList<TextLineInfo> measured)
    {
        var sb = new StringBuilder();
        sb.Append($"상자: {budget.Label}. 한 줄에 약 {budget.CharsPerLine}자가 들어가고, 전체 {budget.Lines}줄 이내여야 한다.\n");
        sb.Append($"지금 {measured.Count}줄이다. PowerPoint가 실제로 나눈 줄:\n");
        int paragraph = 1;
        foreach (var line in measured)
        {
            sb.Append($"  (항목{paragraph}) {line.Text}\n");
            if (line.BreakType is LineBreakType.Enter) paragraph++;
        }
        var orphans = LayoutLint.Segments(measured)
            .Where(s => s.Count >= 2 && s[^1].Text.Trim().Length is > 0 and <= 4)
            .Select(s => s[^1].Text.Trim()).ToList();
        if (orphans.Count > 0) sb.Append($"마지막 줄에 짧게 남은 꼬리: {string.Join(", ", orphans.Select(o => $"'{o}'"))} — 줄이거나 바꿔 앞줄에 붙인다.\n");
        if (measured.Count > budget.Lines)
            sb.Append($"{measured.Count - budget.Lines}줄을 줄여야 한다. 항목을 짧게 하고, 그래도 넘치면 덜 중요한 항목을 뺀다.\n");
        sb.Append("현재 항목(JSON): ").Append(JsonSerializer.Serialize(items, Tools.ToolArgs.Json)).Append('\n');
        sb.Append("고친 항목을 JSON 배열로만 답한다.");
        return sb.ToString();
    }

    /// <summary>모델 답에서 JSON을 꺼낸다. 생각 태그·코드블록·앞뒤 설명을 견딘다.</summary>
    public static string ExtractJson(string text)
    {
        text = Regex.Replace(text, @"<think>.*?</think>", "", RegexOptions.Singleline).Trim();
        int start = text.IndexOfAny(['{', '[']);
        if (start < 0) throw new FormatException("JSON을 찾지 못했습니다.");
        char open = text[start], close = open == '{' ? '}' : ']';
        int end = text.LastIndexOf(close);
        if (end <= start) throw new FormatException("JSON이 닫히지 않았습니다.");
        return text[start..(end + 1)];
    }

    public static IReadOnlyList<CardDraft> ParseDrafts(string text, IReadOnlyList<string> roles)
    {
        using var document = JsonDocument.Parse(ExtractJson(text));
        var root = document.RootElement;
        var cards = root.ValueKind == JsonValueKind.Array ? root : root.GetProperty("cards");
        var drafts = new List<CardDraft>();
        foreach (var card in cards.EnumerateArray())
        {
            string Str(string name) => card.TryGetProperty(name, out var v) ? v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : v.GetRawText() : "";
            var sections = new Dictionary<string, IReadOnlyList<string>>();
            foreach (var role in roles)
                if (card.TryGetProperty(role, out var v))
                    sections[role] = v.ValueKind == JsonValueKind.Array
                        ? v.EnumerateArray().Select(e => e.ValueKind == JsonValueKind.String ? e.GetString() ?? "" : e.GetRawText())
                           .Select(Clean).Where(s => s.Length > 0).ToList()
                        : [Clean(v.GetString() ?? "")];
            drafts.Add(new CardDraft(Str("number").Trim(), Clean(Str("title")), sections));
        }
        return drafts;
    }

    public static IReadOnlyList<string> ParseItems(string text)
    {
        using var document = JsonDocument.Parse(ExtractJson(text));
        var root = document.RootElement;
        // {"items": [...]}처럼 감싸 오면 첫 배열을 쓴다.
        if (root.ValueKind == JsonValueKind.Object)
            foreach (var property in root.EnumerateObject())
                if (property.Value.ValueKind == JsonValueKind.Array) { root = property.Value; break; }
        if (root.ValueKind != JsonValueKind.Array) throw new FormatException("항목 배열이 아닙니다.");
        return root.EnumerateArray().Select(e => Clean(e.ValueKind == JsonValueKind.String ? e.GetString() ?? "" : e.GetRawText()))
                   .Where(s => s.Length > 0).ToList();
    }

    /// <summary>항목 앞의 글머리("-", "•", "1.")와 끝 마침표를 뗀다. 템플릿 문단이 글머리를 이미 갖고 있다.</summary>
    public static string Clean(string item) =>
        Regex.Replace(item.Trim(), @"^(?:[-•·▪◦*]|\d+[.)])\s*", "").TrimEnd('.', '。').Trim();
}

/// <summary>템플릿 카드 하나의 실제 문구. 문체·길이 예시로 문안 작성에 준다.</summary>
public sealed record CardLayoutSample(string Title, IReadOnlyList<(string Label, IReadOnlyList<string> Items)> Sections);
