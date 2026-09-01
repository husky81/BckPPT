namespace CellBox2.BckPPT.Models;

public record RecentPptxFile(
    string   Name,
    string   FullPath,
    DateTime ModifiedAt,
    long     FileSizeBytes)
{
    public string ModifiedText       => ModifiedAt.ToString("MM/dd HH:mm");
    public bool   IsAvailableLocally => FileSizeBytes > 0;

    /// <summary>백그라운드 pre-scan 결과. null = 스캔 전.</summary>
    public int? OverflowCount { get; init; }

    public string StatusIcon => OverflowCount switch
    {
        null => "···",
        0    => "✅",
        _    => $"⚠️ {OverflowCount}개",
    };
}
