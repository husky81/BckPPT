using BckPPT.Infrastructure;

namespace BckPPT.Tests;

public sealed class ExternalCommandValidatorTests
{
    [Fact]
    public void Validate_RejectsOverwriteMode()
    {
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "icon.pptx"));
        var request = new ExternalCommandRequest("1.0", "create_job", Guid.NewGuid().ToString(),
            fixture, "첫 장을 요약해줘", "overwrite");
        Assert.Throws<ArgumentException>(() => ExternalCommandValidator.Validate(request));
    }

    [Fact]
    public void Validate_AcceptsVersionedCopyRequest()
    {
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "icon.pptx"));
        var request = new ExternalCommandRequest("1.0", "create_job", Guid.NewGuid().ToString(),
            fixture, "첫 장을 요약해줘", "copy");
        ExternalCommandValidator.Validate(request);
    }
}
