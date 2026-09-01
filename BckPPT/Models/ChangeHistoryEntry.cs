namespace CellBox2.BckPPT.Models;

public record ChangeHistoryEntry(
    DateTime Timestamp,
    string   PresentationName,
    string   ChangeType,    // "TextReplace" | "AutoFit" | "FontNormalize" | "RuleReplace"
    int      SlideIndex,
    string   ShapeName,
    string   OldValue,
    string   NewValue);
