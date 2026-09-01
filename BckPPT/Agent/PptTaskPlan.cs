using System.Text.Json;
using System.Text.Json.Serialization;

namespace BckPPT.Agent;

public sealed record PptTaskPlan(
    string Summary,
    bool RequiresApproval,
    IReadOnlyList<PptPlanCommand> Commands,
    IReadOnlyList<string> Warnings);

public sealed record PptPlanCommand(
    string Name,
    int? SlideIndex,
    string? ShapeName,
    string? Reason,
    int? SourceSlideIndex = null,
    int? InsertAfterSlideIndex = null,
    string? NewText = null,
    string? ExpectedText = null);

public static class PptTaskPlanDecoder
{
    private static readonly HashSet<string> AllowedCommands = new(StringComparer.Ordinal)
    {
        "analyze_presentation", "search_template_library", "duplicate_slide",
        "replace_shape_text", "insert_slide", "validate_layout", "save_copy"
    };

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    public static PptTaskPlan Decode(string response)
    {
        var json = ExtractJson(response);
        var plan = JsonSerializer.Deserialize<PptTaskPlan>(json, JsonOptions)
            ?? throw new FormatException("빈 작업계획입니다.");
        if (string.IsNullOrWhiteSpace(plan.Summary)) throw new FormatException("작업계획 요약이 없습니다.");
        if (plan.Commands is null || plan.Commands.Count == 0) throw new FormatException("작업 명령이 없습니다.");
        var unknown = plan.Commands.FirstOrDefault(command => !AllowedCommands.Contains(command.Name));
        if (unknown is not null) throw new FormatException($"허용되지 않은 명령입니다: {unknown.Name}");
        if (plan.Commands.Any(command => command.SlideIndex is <= 0))
            throw new FormatException("슬라이드 번호는 1 이상이어야 합니다.");
        return plan with { Warnings = plan.Warnings ?? [] };
    }

    public static string ToDisplayJson(PptTaskPlan plan) => JsonSerializer.Serialize(plan, JsonOptions);

    private static string ExtractJson(string response)
    {
        var text = response.Trim();
        if (text.StartsWith("```"))
        {
            var firstLine = text.IndexOf('\n');
            var lastFence = text.LastIndexOf("```", StringComparison.Ordinal);
            if (firstLine >= 0 && lastFence > firstLine) text = text[(firstLine + 1)..lastFence].Trim();
        }
        var start = text.IndexOf('{');
        var end = text.LastIndexOf('}');
        if (start < 0 || end <= start) throw new FormatException("JSON 작업계획을 찾을 수 없습니다.");
        return text[start..(end + 1)];
    }
}
