using BckPPT.Agent;
using BckPPT.PowerPoint;
using BckPPT.Rag;

namespace BckPPT.Application;

public enum PptJobState { Planned, PreviewReady, Applying, Completed, Cancelled, Failed }

public sealed class PptJob
{
    public required Guid Id { get; init; }
    public required string SourcePath { get; init; }
    public required DateTime SourceLastWriteUtc { get; init; }
    public required string Instruction { get; init; }
    public required PptTaskPlan Plan { get; init; }
    public PptJobState State { get; internal set; } = PptJobState.Planned;
    public string? Preview { get; internal set; }
    public string? ResultPath { get; internal set; }
    public string? Error { get; internal set; }
    public IReadOnlyList<string> BeforeThumbnails { get; internal set; } = [];
    public IReadOnlyList<string> AfterThumbnails { get; internal set; } = [];
    public LayoutValidationResult? Validation { get; internal set; }
    public LayoutDecision? LayoutDecision { get; internal set; }
    public string? ExternalRequestId { get; internal set; }
}
