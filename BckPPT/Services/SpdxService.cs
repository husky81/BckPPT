using System.Text;
using System.Text.RegularExpressions;
using CellBox2.BckPPT.Models;
using CellBox2.BckPPT.ViewModels;

namespace CellBox2.BckPPT.Services;

/// <summary>SPDX(Slide-MPDX) 추출 · 패치 적용 · 검증.</summary>
public class SpdxService
{
    private readonly BckPptViewModel _vm;
    private readonly IPptComService _com;

    // ── 슬라이드별 스냅샷 스택 (undo 지원) ──
    private readonly Dictionary<string, Stack<SlideSnapshot>> _snapshotStacks = new();

    public IPptComService Com => _com;

    public SpdxService(BckPptViewModel vm, IPptComService com)
    {
        _vm = vm;
        _com = com;
    }

    // ── Extract ─────────────────────────────────────────────────

    /// <summary>슬라이드 하나를 SPDX TSV로 직렬화.</summary>
    public (string Tsv, List<SpdxEntry> Entries) ExtractSpdx(string presentationPath, int slideIndex)
    {
        var layout  = _com.GetSlideLayout(presentationPath, slideIndex);
        var entries = new List<SpdxEntry>();
        int nextId  = 1;
        var roleMap = BuildRoleMap(slideIndex);

        // ── 1단계: 모든 shape의 lineInfos 수집 (COM 호출 캐시) ──────
        var regularItems = layout.Shapes
            .Where(s => s.Category == ShapeCategory.Text && s.TextLines?.Count > 0)
            .GroupBy(s => s.Name, StringComparer.Ordinal)
            .Select(g => g.First())
            .Select(s => (
                Name:     s.Name,
                RoleAbbr: roleMap.TryGetValue(s.Name, out var a) ? a : "?",
                IsTable:  false,
                Lines:    _com.GetTextLineBreaks(presentationPath, slideIndex, s.Name)
            ))
            .ToList();

        var tableItems = GetTableCells(slideIndex)
            .Where(s => !string.IsNullOrWhiteSpace(s.FullText))
            .Select(s => (
                Name:     s.ShapeName,
                RoleAbbr: roleMap.TryGetValue(s.ShapeName, out var a) ? a : "표",
                IsTable:  true,
                Lines:    _com.GetTextLineBreaks(_vm.SelectedPresentation!.Path, slideIndex, s.ShapeName)
            ))
            .ToList();

        // ── 2단계: 슬라이드 내 최대 fontSize (상대적 roleHint 계산용) ──
        float slideMaxFont = regularItems.Concat(tableItems)
            .SelectMany(item => item.Lines)
            .Select(l => l.FontSize)
            .DefaultIfEmpty(0f)
            .Max();

        // ── 3단계: entries 구성 ──────────────────────────────────────
        int slideId = nextId++;
        entries.Add(new SpdxEntry(slideId, 0, "slide", 0, 0f, "",
            $"슬라이드 {slideIndex}", slideIndex, "", null));

        foreach (var item in regularItems)
        {
            int shapeId = nextId++;
            entries.Add(new SpdxEntry(shapeId, slideId, $"shape:{item.RoleAbbr}", 0, 0f, "",
                "", slideIndex, item.Name, null));

            for (int li = 0; li < item.Lines.Count; li++)
            {
                var    info     = item.Lines[li];
                int    maxHW    = PptSlideInfo.CalcHalfWidth(info.Text);
                string lineType = info.HasBullet ? "lineB" : "line";
                string roleHint = CalcRoleHint(info.FontSize, slideMaxFont);
                int    lineId   = nextId++;
                entries.Add(new SpdxEntry(lineId, shapeId, lineType, maxHW, info.FontSize, roleHint,
                    info.Text, slideIndex, item.Name, li));
            }
        }

        foreach (var item in tableItems)
        {
            int shapeId = nextId++;
            entries.Add(new SpdxEntry(shapeId, slideId, $"shape:{item.RoleAbbr}", 0, 0f, "",
                "", slideIndex, item.Name, null));

            for (int li = 0; li < item.Lines.Count; li++)
            {
                var    info      = item.Lines[li];
                int    currentHW = PptSlideInfo.CalcHalfWidth(info.Text);
                int    maxHW     = Math.Max(currentHW + 5, 20); // 현재 + 여유 5자
                string lineType  = info.HasBullet ? "lineB" : "line";
                string roleHint  = CalcRoleHint(info.FontSize, slideMaxFont);
                int    lineId    = nextId++;
                entries.Add(new SpdxEntry(lineId, shapeId, lineType, maxHW, info.FontSize, roleHint,
                    info.Text, slideIndex, item.Name, li));
            }
        }

        return (SerializeToTsv(entries, slideIndex), entries);
    }

    // ── Apply ───────────────────────────────────────────────────

    /// <summary>패치 TSV를 파싱하여 적용. 사전·사후 검증 포함.</summary>
    /// <param name="smartTrim">true면 overflow 발생 시 SmartTrim 자동 적용 (최후 수단).</param>
    public SpdxApplyResult ApplyPatch(
        string presentationPath, int slideIndex,
        string patchTsv, List<SpdxEntry> entries,
        bool smartTrim = false)
    {
        var patches = ParsePatch(patchTsv);
        if (patches.Count == 0)
            return new SpdxApplyResult(false, "패치가 비어있습니다.", [], [], []);

        var entryMap = entries.ToDictionary(e => e.Id);
        var violations = new List<string>();
        var changes = new List<SpdxChange>();

        // ── 사전 검증: 노드 존재 + 타입 체크 (maxHW는 경고만, 차단 안함) ──
        var hwWarnings = new List<OverflowEntry>();
        foreach (var (id, newText) in patches)
        {
            if (!entryMap.TryGetValue(id, out var entry))
            {
                violations.Add($"id {id}: 존재하지 않는 노드");
                continue;
            }
            if (!entry.Type.StartsWith("line"))
            {
                violations.Add($"id {id}: line 노드가 아님 (type={entry.Type})");
                continue;
            }

            int hw = PptSlideInfo.CalcHalfWidth(newText);
            if (hw > entry.MaxHW)
                hwWarnings.Add(new OverflowEntry(id, entry.ShapeName, entry.MaxHW, hw, newText));
        }

        if (violations.Count > 0)
            return new SpdxApplyResult(false, "사전 검증 실패", [], violations, []);

        // smartTrim 모드: 사전에 maxHW 초과 텍스트를 축약
        if (smartTrim)
        {
            foreach (var ov in hwWarnings)
                patches[ov.Id] = SmartTrim(patches[ov.Id], ov.MaxHW);
        }

        // ── shape별로 전체 텍스트 재조립 후 적용 ──
        var shapeGroups = entries
            .Where(e => e.Type.StartsWith("line") && e.LineIndex.HasValue)
            .GroupBy(e => e.ShapeName)
            .ToDictionary(g => g.Key, g => g.OrderBy(e => e.LineIndex!.Value).ToList());

        var affectedShapes = patches
            .Select(p => entryMap[p.Key].ShapeName)
            .Distinct()
            .ToList();

        // ── 슬라이드 전체 스냅샷 (undo 스택용) ──
        var slideSnap = CaptureSlideSnapshot(presentationPath, slideIndex, entries);

        // ── 1단계: 모든 shape의 원본 스냅샷 ──
        var snapshots = new Dictionary<string, (string OriginalText, int LineCount)>();
        foreach (var shapeName in affectedShapes)
        {
            var shapeInfo = _com.GetShapeInfo(presentationPath, slideIndex, shapeName);
            if (shapeInfo is not null)
                snapshots[shapeName] = (shapeInfo.FullText, shapeInfo.LineCount);
        }

        // ── 2단계: 모든 shape에 적용 ──
        foreach (var shapeName in affectedShapes)
        {
            if (!shapeGroups.TryGetValue(shapeName, out var lineEntries)) continue;

            var lineBreaks = _com.GetTextLineBreaks(presentationPath, slideIndex, shapeName);
            var newLines = new List<string>();
            for (int i = 0; i < lineEntries.Count; i++)
            {
                var entry = lineEntries[i];
                string text = patches.TryGetValue(entry.Id, out var nt) ? nt : entry.Text;

                if (i < lineBreaks.Count)
                {
                    text += lineBreaks[i].BreakType switch
                    {
                        LineBreakType.Enter => "\r",
                        LineBreakType.ShiftEnter => "\v",
                        _ => ""
                    };
                }
                newLines.Add(text);
            }

            string fullText = string.Join("", newLines);
            string originalText = snapshots.TryGetValue(shapeName, out var s) ? s.OriginalText : "";
            _com.SetShapeText(presentationPath, slideIndex, shapeName, fullText);
            changes.Add(new SpdxChange(shapeName, originalText, fullText));
        }

        // ── 3단계: 사후 검증 (lineCount) ──
        var overflowEntries = new List<OverflowEntry>();
        foreach (var shapeName in affectedShapes)
        {
            if (!snapshots.TryGetValue(shapeName, out var snap)) continue;
            var after = _com.GetShapeInfo(presentationPath, slideIndex, shapeName);
            if (after is not null && after.LineCount != snap.LineCount)
            {
                violations.Add(
                    $"{shapeName}: lineCount {snap.LineCount}→{after.LineCount} (overflow)");

                // overflow 발생한 shape의 패치된 줄 정보 수집 (LLM 재요청용)
                if (shapeGroups.TryGetValue(shapeName, out var lineEntries))
                {
                    foreach (var entry in lineEntries)
                    {
                        if (patches.TryGetValue(entry.Id, out var patchedText))
                        {
                            int hw = PptSlideInfo.CalcHalfWidth(patchedText);
                            overflowEntries.Add(new OverflowEntry(
                                entry.Id, shapeName, entry.MaxHW, hw, patchedText));
                        }
                    }
                }
            }
        }

        // ── 4단계: 위반 있으면 전체 롤백 ──
        if (violations.Count > 0)
        {
            foreach (var change in changes)
                _com.SetShapeText(presentationPath, slideIndex, change.ShapeName, change.OriginalText);
            return new SpdxApplyResult(false, "사후 검증 실패 — 전체 롤백됨", changes, violations, overflowEntries);
        }

        // 성공: 스냅샷을 undo 스택에 push
        PushSnapshot(presentationPath, slideIndex, slideSnap);

        return new SpdxApplyResult(true, $"{changes.Count}개 shape 적용 완료", changes, [], []);
    }

    // ── Undo ────────────────────────────────────────────────────

    /// <summary>슬라이드 전체를 마지막 스냅샷으로 복원.</summary>
    public SpdxUndoResult Undo(string presentationPath, int slideIndex)
    {
        string key = SnapshotKey(presentationPath, slideIndex);
        if (!_snapshotStacks.TryGetValue(key, out var stack) || stack.Count == 0)
            return new SpdxUndoResult(false, "복원할 스냅샷이 없습니다.", 0);

        var snapshot = stack.Pop();
        int restored = 0;
        foreach (var (shapeName, text) in snapshot.ShapeTexts)
        {
            _com.SetShapeText(presentationPath, slideIndex, shapeName, text);
            restored++;
        }

        return new SpdxUndoResult(true,
            $"슬라이드 {slideIndex}을 {snapshot.CreatedAt:HH:mm:ss} 스냅샷으로 복원 ({restored}개 shape)",
            stack.Count);
    }

    /// <summary>슬라이드의 남은 undo 스택 수를 반환.</summary>
    public int GetUndoCount(string presentationPath, int slideIndex)
    {
        string key = SnapshotKey(presentationPath, slideIndex);
        return _snapshotStacks.TryGetValue(key, out var stack) ? stack.Count : 0;
    }

    // ── SmartTrim ───────────────────────────────────────────────

    /// <summary>최후 수단 텍스트 축약: 괄호 제거 → 수식어 축약 → 말줄임.</summary>
    internal static string SmartTrim(string text, int maxHW)
    {
        if (PptSlideInfo.CalcHalfWidth(text) <= maxHW) return text;

        // 1순위: 괄호 부분 제거 "(출처, 2024)" "（참고）" etc.
        string trimmed = Regex.Replace(text, @"\s*[\(（\[]\s*[^)）\]]*[\)）\]]", "").Trim();
        if (PptSlideInfo.CalcHalfWidth(trimmed) <= maxHW) return trimmed;

        // 2순위: 한국어 수식어 축약
        var suffixPatterns = new[]
        {
            (@"할\s*것으로\s*전망", "전망"),
            (@"하는\s*것으로\s*나타남", "나타남"),
            (@"되는\s*것으로\s*보임", "보임"),
            (@"할\s*것으로\s*예상", "예상"),
            (@"인\s*것으로\s*분석", "분석"),
            (@"할\s*필요가\s*있음", "필요"),
        };
        foreach (var (pattern, replacement) in suffixPatterns)
        {
            trimmed = Regex.Replace(trimmed, pattern, replacement);
            if (PptSlideInfo.CalcHalfWidth(trimmed) <= maxHW) return trimmed;
        }

        // 3순위: 끝에서 말줄임 (글자 단위로 제거)
        while (PptSlideInfo.CalcHalfWidth(trimmed) > maxHW && trimmed.Length > 3)
            trimmed = trimmed[..^1];
        return trimmed + "…";
    }

    // ── Snapshot helpers ────────────────────────────────────────

    private SlideSnapshot CaptureSlideSnapshot(
        string presentationPath, int slideIndex, List<SpdxEntry> entries)
    {
        var shapeTexts = new Dictionary<string, string>();
        var shapeNames = entries
            .Where(e => e.Type.StartsWith("shape:"))
            .Select(e => e.ShapeName)
            .Distinct();

        foreach (var shapeName in shapeNames)
        {
            var info = _com.GetShapeInfo(presentationPath, slideIndex, shapeName);
            if (info is not null)
                shapeTexts[shapeName] = info.FullText;
        }

        return new SlideSnapshot(slideIndex, shapeTexts, DateTime.Now);
    }

    private void PushSnapshot(string presentationPath, int slideIndex, SlideSnapshot snapshot)
    {
        string key = SnapshotKey(presentationPath, slideIndex);
        if (!_snapshotStacks.TryGetValue(key, out var stack))
        {
            stack = new Stack<SlideSnapshot>();
            _snapshotStacks[key] = stack;
        }

        // 최대 10개 스냅샷 유지
        if (stack.Count >= 10)
        {
            var list = stack.ToList();
            list.RemoveAt(list.Count - 1);
            stack.Clear();
            for (int i = list.Count - 1; i >= 0; i--)
                stack.Push(list[i]);
        }

        stack.Push(snapshot);
    }

    private static string SnapshotKey(string presentationPath, int slideIndex)
        => $"{presentationPath}::{slideIndex}";

    // ── Helpers ──────────────────────────────────────────────────

    private Dictionary<string, string> BuildRoleMap(int slideIndex)
    {
        var map = new Dictionary<string, string>();
        foreach (var node in FlattenNodes(_vm.RootNodes))
        {
            if (node.Shape.SlideIndex == slideIndex)
                map[node.Shape.ShapeName] = node.Shape.Role.ToAbbr();
        }
        return map;
    }

    private IEnumerable<PptSlideInfo> GetTableCells(int slideIndex)
    {
        // _vm의 CurrentShapes/AllShapes에서 표 셀 (shapeName에 [R C] 포함) 필터
        return FlattenNodes(_vm.RootNodes)
            .Where(n => n.Shape.SlideIndex == slideIndex
                     && n.Shape.ShapeName.Contains(" [R"))
            .Select(n => n.Shape);
    }

    private static string SerializeToTsv(List<SpdxEntry> entries, int slideIndex)
    {
        int shapeCount = entries.Count(e => e.Type.StartsWith("shape:"));
        var sb = new StringBuilder();
        sb.AppendLine($"# SPDX v1.0 — Slide {slideIndex}, {shapeCount} shapes");
        sb.AppendLine("# id\tparent\ttype\tmaxHW\tfontSize\troleHint\ttext");
        sb.AppendLine("# AI patch: id\\tnewText (한 줄씩, 변경된 줄만)");
        foreach (var e in entries)
        {
            string fs = e.FontSize > 0 ? e.FontSize.ToString("0") : "";
            sb.AppendLine($"{e.Id}\t{e.ParentId}\t{e.Type}\t{e.MaxHW}\t{fs}\t{e.RoleHint}\t{e.Text}");
        }
        return sb.ToString();
    }

    /// <summary>슬라이드 내 최대 fontSize 기준 상대값으로 roleHint 계산.
    /// slideMaxFont <= 0 이면 절대값 기준 fallback.</summary>
    private static string CalcRoleHint(float fontSize, float slideMaxFont)
    {
        if (fontSize <= 0f) return "";
        if (slideMaxFont > 0f)
        {
            if (fontSize >= slideMaxFont - 1f) return "title";
            if (fontSize >= slideMaxFont - 3f) return "subtitle";
            return "body";
        }
        // fallback: 절대값 기준
        return fontSize >= 14f ? "title" : fontSize >= 11f ? "subtitle" : "body";
    }

    private static Dictionary<int, string> ParsePatch(string patchTsv)
    {
        var result = new Dictionary<int, string>();
        foreach (var raw in patchTsv.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            // 자동 줄바꿈 조각 끝의 공백은 다음 조각과 단어를 구분하는
            // 의미 있는 문자이므로 제거하지 않는다. CR만 정리한다.
            var line = raw.TrimEnd('\r');
            if (line.TrimStart().StartsWith('#')) continue;
            var parts = line.Split('\t', 2);
            if (parts.Length == 2 && int.TryParse(parts[0], out int id))
                result[id] = parts[1];
        }
        return result;
    }

    private static IEnumerable<PptNode> FlattenNodes(IEnumerable<PptNode> nodes)
    {
        foreach (var n in nodes)
        {
            yield return n;
            foreach (var child in FlattenNodes(n.Children))
                yield return child;
        }
    }
}

// ── Result types ────────────────────────────────────────────────

public record SpdxApplyResult(
    bool Ok,
    string Message,
    List<SpdxChange> Changes,
    List<string> Violations,
    List<OverflowEntry> OverflowEntries);

public record SpdxChange(
    string ShapeName,
    string OriginalText,
    string NewText);

/// <summary>overflow 발생 줄 정보 — LLM 재요청 시 "이 줄을 더 짧게" 프롬프트용.</summary>
public record OverflowEntry(
    int Id,
    string ShapeName,
    int MaxHW,
    int ActualHW,
    string Text);

/// <summary>슬라이드 전체 텍스트 스냅샷 (undo용).</summary>
public record SlideSnapshot(
    int SlideIndex,
    Dictionary<string, string> ShapeTexts,
    DateTime CreatedAt);

public record SpdxUndoResult(
    bool Ok,
    string Message,
    int RemainingUndoCount);
