using System.Text.RegularExpressions;

namespace BckPPT.Models;

/// <summary>
/// 슬라이드 안 도형을 가리키는 주소. 기존 "도형 이름" 규약을 그대로 받으면서
/// 같은 이름이 여러 개인 경우(복사해 만든 카드 등)와 그룹 안 도형을 구분한다.
/// <list type="bullet">
/// <item><c>Text 19</c> — 그룹 안까지 펼친 순서에서 첫 번째 "Text 19"</item>
/// <item><c>Text 19#3</c> — 같은 이름 중 세 번째(1부터)</item>
/// <item><c>Group 171/Text 19</c> — 바로 위 그룹이 "Group 171"인 것만</item>
/// <item><c>Table 1 [R2C3]</c> — 표 셀. 위 형식과 함께 쓸 수 있다.</item>
/// </list>
/// 이름 자체에 '/'나 '#'이 들어간 도형은 전체 문자열이 이름과 정확히 같으면 그 도형을 먼저 고른다.
/// </summary>
public sealed record ShapeAddress(
    string                Raw,
    IReadOnlyList<string> GroupPath,
    string                Name,
    int                   Occurrence,
    int?                  Row = null,
    int?                  Col = null)
{
    private static readonly Regex CellSuffix  = new(@"^(.+?)\s+\[R(\d+)C(\d+)\]$");
    private static readonly Regex Occurrence_ = new(@"^(.+)#(\d+)$");

    public bool IsTableCell => Row is not null && Col is not null;

    public static ShapeAddress Parse(string address)
    {
        string rest = address.Trim();
        int? row = null, col = null;

        var cell = CellSuffix.Match(rest);
        if (cell.Success)
        {
            rest = cell.Groups[1].Value;
            row  = int.Parse(cell.Groups[2].Value);
            col  = int.Parse(cell.Groups[3].Value);
        }

        var segments = rest.Split('/');
        string last  = segments[^1];
        int occurrence = 1;
        var occ = Occurrence_.Match(last);
        if (occ.Success && int.Parse(occ.Groups[2].Value) >= 1)
        {
            last       = occ.Groups[1].Value;
            occurrence = int.Parse(occ.Groups[2].Value);
        }

        return new ShapeAddress(address, segments[..^1], last, occurrence, row, col);
    }

    /// <summary>
    /// 후보(펼친 순서의 도형 이름과 그 조상 그룹 경로) 중 주소가 가리키는 위치를 돌려준다. 없으면 -1.
    /// </summary>
    public int Select(IReadOnlyList<(IReadOnlyList<string> GroupPath, string Name)> candidates)
    {
        // 표 셀 접미사를 뗀 나머지가 이름과 정확히 같으면 리터럴 이름으로 본다.
        string literal = IsTableCell ? CellSuffix.Match(Raw.Trim()).Groups[1].Value : Raw.Trim();
        for (int i = 0; i < candidates.Count; i++)
            if (candidates[i].Name == literal) return i;

        int seen = 0;
        for (int i = 0; i < candidates.Count; i++)
        {
            var (path, name) = candidates[i];
            if (name != Name || !EndsWith(path, GroupPath)) continue;
            if (++seen == Occurrence) return i;
        }
        return -1;
    }

    /// <summary>목록에 보여줄 주소. 이름이 하나뿐이면 이름만, 겹치면 "이름#n".</summary>
    public static string Format(string name, int occurrence, int sameNameCount)
        => sameNameCount > 1 ? $"{name}#{occurrence}" : name;

    private static bool EndsWith(IReadOnlyList<string> path, IReadOnlyList<string> suffix)
    {
        if (suffix.Count > path.Count) return false;
        for (int i = 1; i <= suffix.Count; i++)
            if (path[^i] != suffix[^i]) return false;
        return true;
    }
}
