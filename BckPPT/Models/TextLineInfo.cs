namespace CellBox2.BckPPT.Models;

public enum LineBreakType
{
    /// <summary>마지막 줄 (줄넘김 없음)</summary>
    None,
    /// <summary>텍스트가 상자 폭을 초과하여 PPT가 자동으로 줄바꿈</summary>
    AutoWrap,
    /// <summary>사용자가 Enter로 줄바꿈 (새 문단)</summary>
    Enter,
    /// <summary>사용자가 Shift+Enter로 줄바꿈 (문단 유지)</summary>
    ShiftEnter
}

public record TextLineInfo(
    int LineNumber,
    string Text,
    LineBreakType BreakType,
    string BulletText = "",
    int IndentLevel = 0,
    float FontSize = 12f)
{
    /// <summary>문단의 첫 줄인 경우 불릿 표시.</summary>
    public bool HasBullet => !string.IsNullOrEmpty(BulletText);

    /// <summary>인덴트에 따른 왼쪽 여백 (px).</summary>
    public double IndentMargin => IndentLevel > 1 ? (IndentLevel - 1) * 16.0 : 0;

    public string BreakIcon => BreakType switch
    {
        LineBreakType.AutoWrap   => "\u2192",   // →
        LineBreakType.Enter      => "\u21B5",   // ↵
        LineBreakType.ShiftEnter => "\u2193",   // ↓
        _                        => ""
    };

    public string BreakLabel => BreakType switch
    {
        LineBreakType.AutoWrap   => "auto",
        LineBreakType.Enter      => "Enter",
        LineBreakType.ShiftEnter => "Shift+Enter",
        _                        => ""
    };
}
