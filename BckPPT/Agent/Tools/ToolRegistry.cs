using System.Text.Json;
using BckPPT.Llm;

namespace BckPPT.Agent.Tools;

public sealed record ToolOutcome(string Content, bool IsError);

/// <summary>이름으로 도구를 찾아 실행한다. 모델이 고칠 수 있는 실패는 예외 대신 오류 문장으로 돌려준다.</summary>
public sealed class ToolRegistry
{
    private readonly Dictionary<string, AgentTool> _tools = new(StringComparer.Ordinal);

    public ToolRegistry(IEnumerable<AgentTool> tools)
    {
        foreach (var tool in tools) _tools[tool.Name] = tool;
    }

    public IReadOnlyCollection<AgentTool> Tools => _tools.Values;

    public IReadOnlyList<LlmToolDefinition> Definitions() => _tools.Values.Select(t => t.ToDefinition()).ToList();

    public bool TryGet(string name, out AgentTool tool) => _tools.TryGetValue(name, out tool!);

    public async Task<ToolOutcome> ExecuteAsync(AgentToolContext context, LlmToolCall call, CancellationToken cancellationToken)
    {
        if (!_tools.TryGetValue(call.Name, out var tool))
            return new($"오류: '{call.Name}' 도구는 없습니다. 사용 가능: {string.Join(", ", _tools.Keys)}", true);

        JsonDocument arguments;
        try { arguments = JsonDocument.Parse(string.IsNullOrWhiteSpace(call.ArgumentsJson) ? "{}" : call.ArgumentsJson); }
        catch (JsonException ex) { return new($"오류: 인자가 올바른 JSON이 아닙니다 — {ex.Message}", true); }

        using (arguments)
        {
            if (arguments.RootElement.ValueKind != JsonValueKind.Object)
                return new("오류: 인자는 JSON 객체여야 합니다.", true);
            try
            {
                return new(await tool.Execute(context, arguments.RootElement, cancellationToken), false);
            }
            catch (OperationCanceledException) { throw; }
            catch (AgentToolException ex) { return new($"오류: {ex.Message}", true); }
            catch (Exception ex)
            {
                // COM 오류(잘못된 슬라이드 번호 등)도 모델이 다음 행동을 고를 수 있게 알린다.
                return new($"오류: {call.Name} 실행 실패 — {ex.GetBaseException().Message}", true);
            }
        }
    }
}
