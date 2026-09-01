using BckPPT.Agent;

namespace BckPPT.Tests;

public sealed class PptTaskPlanDecoderTests
{
    [Fact]
    public void Decode_RejectsUnknownCommand()
    {
        const string json = """{"summary":"bad","requiresApproval":true,"commands":[{"name":"run_code"}],"warnings":[]}""";
        Assert.Throws<FormatException>(() => PptTaskPlanDecoder.Decode(json));
    }

    [Fact]
    public void Decode_AcceptsFencedAllowedPlan()
    {
        const string json = """
            ```json
            {"summary":"ok","requiresApproval":true,"commands":[{"name":"analyze_presentation","slideIndex":1}],"warnings":[]}
            ```
            """;
        var plan = PptTaskPlanDecoder.Decode(json);
        Assert.Equal("analyze_presentation", plan.Commands[0].Name);
    }
}
