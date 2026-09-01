namespace BckPPT.Models;

public record PptSlideInfo(
    int    SlideIndex,
    string ShapeName,
    string FullText,
    string TextPreview,
    float  BoundHeight,
    float  BoxHeight,
    int    LineCount,
    bool   IsOverflow,
    float  FontSize = 0f)
{
    /// <summary>빈 텍스트 상자 여부. 글꼴 필터 시에만 표시.</summary>
    public bool IsEmpty { get; init; }

    /// <summary>이미지 shape 여부.</summary>
    public bool IsImage { get; init; }

    /// <summary>줄 중 최대 반각 글자 수 (전각=2, 반각=1). 분석 시 채워짐. LLM 텍스트 생성 상한으로 사용.</summary>
    public int MaxLineHw { get; init; }

    /// <summary>하나 이상의 줄에 불릿이 있으면 true. 분석 시 채워짐.</summary>
    public bool HasBullet { get; init; }

    /// <summary>분류된 역할.</summary>
    public ShapeRole Role { get; set; } = ShapeRole.Unknown;

    /// <summary>사용자가 직접 라벨링한 역할인지.</summary>
    public bool IsUserLabeled { get; set; }

    /// <summary>역할 약어 (표시용).</summary>
    public string RoleAbbr => Role.ToAbbr();

    public string OverflowIcon => IsOverflow ? "\u26a0\ufe0f" : "\u2705";

    /// <summary>반각 기준 글자 수 (전각=2, 반각=1).</summary>
    public int HalfWidthLength => CalcHalfWidth(FullText);

    internal static int CalcHalfWidth(string s)
    {
        int len = 0;
        foreach (char c in s)
        {
            if (c == '\r' || c == '\n') continue;
            // CJK, 전각 문자 → 2, 그 외 → 1
            len += (c >= 0x1100 && c <= 0x115F) || // 한글 자모
                   (c >= 0x2E80 && c <= 0x9FFF) || // CJK
                   (c >= 0xAC00 && c <= 0xD7AF) || // 한글 음절
                   (c >= 0xF900 && c <= 0xFAFF) || // CJK 호환
                   (c >= 0xFF01 && c <= 0xFF60) || // 전각 영숫자
                   (c >= 0xFFE0 && c <= 0xFFE6)    // 전각 기호
                   ? 2 : 1;
        }
        return len;
    }

    public float LineHeight => LineCount > 0
        ? (float)Math.Round(BoundHeight / LineCount, 1)
        : 0f;

    public float OverflowAmount => IsOverflow
        ? (float)Math.Round(BoundHeight - BoxHeight, 1)
        : 0f;
}
