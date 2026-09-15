using System.IO;

namespace BckPPT.Application;

/// <summary>
/// 명령줄 인자 해석 결과.
/// CellBox 트레이가 플러그인을 실행할 때 <c>--sync-folder "&lt;경로&gt;"</c> 로
/// 동기화 폴더를 넘기며, 이 값은 파일 열기 대화상자의 기본 경로로 쓰인다.
/// </summary>
public sealed record StartupOptions(string? SyncFolder)
{
    public static StartupOptions Empty { get; } = new(SyncFolder: null);

    private const string SyncFolderFlag = "--sync-folder";

    public static StartupOptions Parse(IReadOnlyList<string>? args)
    {
        if (args is null) return Empty;

        for (var index = 0; index < args.Count; index++)
        {
            var argument = args[index];
            if (argument is null) continue;

            // --sync-folder=<경로>
            if (argument.StartsWith(SyncFolderFlag + "=", StringComparison.OrdinalIgnoreCase))
                return FromRaw(argument[(SyncFolderFlag.Length + 1)..]);

            // --sync-folder <경로>
            if (argument.Equals(SyncFolderFlag, StringComparison.OrdinalIgnoreCase) && index + 1 < args.Count)
                return FromRaw(args[index + 1]);
        }

        return Empty;
    }

    private static StartupOptions FromRaw(string? raw)
    {
        var value = raw?.Trim().Trim('"');
        return string.IsNullOrWhiteSpace(value) ? Empty : new StartupOptions(value);
    }

    /// <summary>존재하는 폴더일 때만 대화상자 기본 경로로 쓴다. 없는 경로는 무시한다.</summary>
    public string? ResolveInitialDirectory()
    {
        if (string.IsNullOrWhiteSpace(SyncFolder)) return null;
        try { return Directory.Exists(SyncFolder) ? Path.GetFullPath(SyncFolder) : null; }
        catch { return null; }
    }
}
