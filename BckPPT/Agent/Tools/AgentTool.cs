using System.Text.Encodings.Web;
using System.Text.Json;
using BckPPT.Llm;
using BckPPT.Services;

namespace BckPPT.Agent.Tools;

/// <summary>
/// 모델이 부를 수 있는 도구 하나. 결과는 모델에게 그대로 보여줄 문자열이다.
/// <see cref="Writes"/>가 참인 도구는 작업 사본을 바꾼다.
/// </summary>
public sealed record AgentTool(
    string Name,
    string Description,
    string ParametersSchema,
    bool Writes,
    Func<AgentToolContext, JsonElement, CancellationToken, Task<string>> Execute)
{
    public LlmToolDefinition ToDefinition()
    {
        using var schema = JsonDocument.Parse(ParametersSchema);
        return new LlmToolDefinition(Name, Description, schema.RootElement.Clone());
    }
}

/// <summary>모델에게 돌려줄 오류. 인자가 틀렸거나 도형을 못 찾은 경우처럼 모델이 고칠 수 있는 것.</summary>
public sealed class AgentToolException(string message) : Exception(message);

/// <summary>
/// 도구 실행 환경. 모든 쓰기는 <see cref="WorkingPath"/>(작업 사본)에만 한다.
/// 원본 경로는 도구에 노출하지 않는다.
/// </summary>
public sealed class AgentToolContext(IPptComService ppt, string workingPath)
{
    public IPptComService Ppt { get; } = ppt;
    public string WorkingPath { get; } = workingPath;
    public SlideOrigins Origins { get; } = new();

    /// <summary>
    /// 판정 기준으로 읽기만 하는 원본 파일. 작업 사본의 슬라이드를 원본의 같은 슬라이드(복제본이면 복제 원본)와 비교한다.
    /// 도구 인자로는 노출하지 않는다. 없으면 작업 사본 안의 복제 원본을 기준으로 쓴다.
    /// </summary>
    public string? BaselinePath { get; init; }

    /// <summary>세션 중 추가한 도형 (슬라이드, 주소). 겹침 판정 대상.</summary>
    public HashSet<(int Slide, string Address)> AddedShapes { get; } = new();

    /// <summary>사용자에게 보여줄 변경 목록.</summary>
    public List<string> Changes { get; } = new();

    /// <summary>이미지 경로와 질문을 받아 비전 모델의 설명을 돌려준다. 없으면 렌더 결과는 경로만 알린다.</summary>
    public Func<string, string, CancellationToken, Task<string>>? DescribeImage { get; init; }

    /// <summary>
    /// (지시, 원문 발췌)를 받아 별도 LLM 호출로 요약을 돌려준다. 긴 원문이 주 대화의 32k 문맥을 채우지 않게 한다.
    /// 없으면 brief_source는 발췌를 그대로 돌려준다.
    /// </summary>
    public Func<string, string, CancellationToken, Task<string>>? Summarize { get; init; }

    /// <summary>카드 문안처럼 품질이 중요한 글을 쓰는 호출(GITOKY_LLM_WRITER_MODEL). 없으면 Summarize를 쓴다.</summary>
    public Func<string, string, CancellationToken, Task<string>>? Write { get; init; }

    /// <summary>진행 상황(파이프라인 단계)을 알리는 곳. 채팅창 진행 로그로 간다.</summary>
    public IProgress<string>? Status { get; set; }

    /// <summary>모델별 사용량. 파이프라인 결과에 함께 보고한다.</summary>
    public BckPPT.Llm.UsageMeter? Meter { get; set; }

    /// <summary>원문 라이브러리. Phase D 전까지는 비어 있을 수 있다.</summary>
    public ISourceLibrary? Sources { get; set; }

    public string? FinishSummary { get; set; }
}

/// <summary>원문 검색. 쪽 단위 조각을 돌려준다.</summary>
public interface ISourceLibrary
{
    IReadOnlyList<SourceHit> Search(string query, int limit);

    /// <summary>문서 목록(document가 null) 또는 한 문서의 절 제목과 쪽 범위.</summary>
    string Outline(string? document);

    /// <summary>한 문서에서 여러 질의의 상위 쪽 발췌(쪽 순서, 최대 글자 수).</summary>
    string Excerpts(string document, IReadOnlyList<string> queries, int pagesPerQuery, int maxChars);
    string Read(string document, int fromPage, int toPage);
}

public sealed record SourceHit(string Document, int Page, string Snippet);

/// <summary>숫자와 + - * / ( )만 있는 식을 계산한다. 그 밖의 글자가 있으면 거부한다.</summary>
public static class Arithmetic
{
    private static readonly System.Text.RegularExpressions.Regex Allowed = new(@"^[\d\s.+\-*/()]+$");

    public static bool TryEvaluate(string expression, out float value)
    {
        value = 0;
        if (!Allowed.IsMatch(expression)) return false;
        try
        {
            var result = new System.Data.DataTable().Compute(expression, null);
            value = Convert.ToSingle(result, System.Globalization.CultureInfo.InvariantCulture);
            return float.IsFinite(value);
        }
        catch { return false; }
    }
}

/// <summary>도구 인자 읽기. 빠지거나 형식이 틀리면 모델이 고칠 수 있는 문장으로 알린다.</summary>
public static class ToolArgs
{
    public static readonly JsonSerializerOptions Json = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    public static string ToJson(object value) => JsonSerializer.Serialize(value, Json);

    public static int Int(JsonElement args, string name) =>
        OptionalInt(args, name) ?? throw new AgentToolException($"인자 '{name}'(정수)가 필요합니다.");

    public static int? OptionalInt(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var v) || v.ValueKind == JsonValueKind.Null) return null;
        if (v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out var i)) return i;
        if (v.ValueKind == JsonValueKind.String && int.TryParse(v.GetString(), out i)) return i;
        throw new AgentToolException($"인자 '{name}'는 정수여야 합니다.");
    }

    public static float? OptionalFloat(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var v) || v.ValueKind == JsonValueKind.Null) return null;
        if (v.ValueKind == JsonValueKind.Number) return v.GetSingle();
        if (v.ValueKind == JsonValueKind.String &&
            float.TryParse(v.GetString(), System.Globalization.CultureInfo.InvariantCulture, out var f)) return f;
        // 모델이 "216*4"처럼 계산식을 보내기도 한다. 숫자와 사칙연산만 있으면 계산해 준다.
        if (v.ValueKind == JsonValueKind.String && Arithmetic.TryEvaluate(v.GetString()!, out var computed)) return computed;
        throw new AgentToolException($"인자 '{name}'는 숫자여야 합니다.");
    }

    public static bool? OptionalBool(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var v) || v.ValueKind == JsonValueKind.Null) return null;
        return v.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String when bool.TryParse(v.GetString(), out var b) => b,
            _ => throw new AgentToolException($"인자 '{name}'는 true/false여야 합니다.")
        };
    }

    public static string String(JsonElement args, string name) =>
        OptionalString(args, name) ?? throw new AgentToolException($"인자 '{name}'(문자열)가 필요합니다.");

    public static string? OptionalString(JsonElement args, string name) =>
        args.TryGetProperty(name, out var v) && v.ValueKind != JsonValueKind.Null
            ? v.ValueKind == JsonValueKind.String ? v.GetString() : v.GetRawText()
            : null;

    public static IReadOnlyList<JsonElement> ObjectArray(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var v)) throw new AgentToolException($"인자 '{name}'(객체 배열)가 필요합니다.");
        // 모델이 배열을 JSON 문자열로 한 번 더 감싸 보내기도 한다.
        if (v.ValueKind == JsonValueKind.String)
        {
            try { v = JsonDocument.Parse(v.GetString()!).RootElement.Clone(); }
            catch (JsonException) { throw new AgentToolException($"인자 '{name}'는 객체 배열이어야 합니다."); }
        }
        if (v.ValueKind != JsonValueKind.Array) throw new AgentToolException($"인자 '{name}'는 객체 배열이어야 합니다.");
        var items = v.EnumerateArray().Where(e => e.ValueKind == JsonValueKind.Object).ToList();
        if (items.Count == 0) throw new AgentToolException($"인자 '{name}'가 비었습니다.");
        return items;
    }

    public static IReadOnlyList<string> StringArray(JsonElement args, string name)
    {
        if (!args.TryGetProperty(name, out var v) || v.ValueKind != JsonValueKind.Array)
            throw new AgentToolException($"인자 '{name}'(문자열 배열)가 필요합니다.");
        return v.EnumerateArray()
            .Select(e => e.ValueKind == JsonValueKind.String ? e.GetString() ?? "" : e.GetRawText())
            .ToList();
    }
}
