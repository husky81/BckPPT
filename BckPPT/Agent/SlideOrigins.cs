namespace BckPPT.Agent;

/// <summary>
/// 세션 중 복제로 생긴 슬라이드가 어느 슬라이드에서 왔는지 기억한다. LayoutLint의 기준(baseline)이 된다.
/// 슬라이드를 끼워 넣으면 뒤 번호가 밀리므로 기록도 함께 민다. 복제본을 다시 복제해도 기준은 처음 원본이다.
/// </summary>
public sealed class SlideOrigins
{
    private Dictionary<int, int> _origins = new();
    private HashSet<int> _touched = new();

    public IReadOnlyDictionary<int, int> Map => _origins;

    /// <summary>쓰기 도구가 건드린 슬라이드. 루프 종료 전 Lint 대상.</summary>
    public IReadOnlyCollection<int> Touched => _touched;

    public int? BaselineOf(int slide) => _origins.TryGetValue(slide, out var origin) ? origin : null;

    public void MarkTouched(int slide) => _touched.Add(slide);

    /// <summary><paramref name="source"/>를 복제해 <paramref name="newIndex"/>에 생겼다.</summary>
    public void RecordDuplicate(int source, int newIndex)
    {
        int Shift(int slide) => slide >= newIndex ? slide + 1 : slide;
        _origins = _origins.ToDictionary(kv => Shift(kv.Key), kv => Shift(kv.Value));
        _touched = _touched.Select(Shift).ToHashSet();
        int shiftedSource = Shift(source);
        _origins[newIndex] = _origins.GetValueOrDefault(shiftedSource, shiftedSource);
        _touched.Add(newIndex);
    }
}
