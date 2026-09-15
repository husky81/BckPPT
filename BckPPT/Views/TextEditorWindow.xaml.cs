using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Views;

/// <summary>
/// 문단마다 ID를 붙여 보여주고, 수정된 내용을 다시 받아들이는 창.
///
/// 입력은 두 가지를 모두 받는다.
///  1) 목록을 이 창에서 직접 고치기
///  2) 외부 에이전트의 답변을 통째로 붙여넣기
/// 둘 다 같은 파서를 타므로 형식을 맞출 필요가 없다.
/// </summary>
public partial class TextEditorWindow : Window
{
    private readonly IReadOnlyList<ParagraphEntry>                 _entries;
    private readonly Dictionary<(int, string), PptSlideInfo>       _shapeLookup;
    private readonly string                                        _originalBody;

    /// <summary>사용자가 확인한 변경 목록. 취소하면 null.</summary>
    public List<ParagraphChange>? AppliedChanges { get; private set; }

    /// <summary>치환한 문단을 빨간색으로 표시할지.</summary>
    public bool MarkRed { get; private set; } = true;

    private const string Prompt =
        "# 아래는 PPT의 문단별 텍스트입니다. 요청에 맞게 수정해 주세요.\n" +
        "- 수정이 필요한 항목만 \"ID: 새 텍스트\" 형식으로 답해 주세요.\n" +
        "- ID는 반드시 원본 그대로 유지해 주세요. (예: S1-12-3)\n" +
        "- 문단을 지우려면 \"S1-12-3: (빈 문단)\" 으로 써 주세요.\n" +
        "- 문단을 추가하려면 \"S1-12-3+: 새 문단\" 처럼 ID 뒤에 +를 붙여 주세요.\n" +
        "- 문단 안에서 줄을 나누려면 \\n 으로 표시해 주세요.\n" +
        "- 설명을 함께 써 주셔도 됩니다. ID가 붙은 줄만 반영됩니다.";

    public TextEditorWindow(
        IReadOnlyList<ParagraphEntry> entries,
        Dictionary<(int, string), PptSlideInfo> shapeLookup,
        string slideTitle)
    {
        InitializeComponent();
        _entries     = entries;
        _shapeLookup = shapeLookup;
        Title        = $"텍스트 치환 — {slideTitle}";

        SourceInitialized += (_, _) =>
        {
            var hwnd  = new WindowInteropHelper(this).Handle;
            int value = 1;
            DwmSetWindowAttribute(hwnd, 20, ref value, sizeof(int));
        };

        PromptBox.Text = Prompt;

        var sb = new StringBuilder();
        int lastShapeSeq = -1;
        foreach (var entry in _entries)
        {
            // 도형이 바뀔 때 빈 줄을 넣어 사람이 덩어리를 알아보게 한다.
            if (lastShapeSeq >= 0 && entry.ShapeSeq != lastShapeSeq) sb.AppendLine();
            lastShapeSeq = entry.ShapeSeq;
            sb.AppendLine($"{entry.Id}: {Escape(entry.Text)}");
        }
        _originalBody = sb.ToString();
        Editor.Text   = _originalBody;
    }

    /// <summary>문단 안의 Shift+Enter를 한 줄로 보이게. 파서가 다시 \v로 되돌린다.</summary>
    private static string Escape(string s) => s.Replace("\v", "\\n");

    // ── 반영 ─────────────────────────────────────────────────────

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        var parsed  = PastedPlanParser.Parse(Editor.Text, _entries);
        var byId    = _entries.ToDictionary(x => x.Id, StringComparer.OrdinalIgnoreCase);
        var changes = new List<ParagraphChange>();

        foreach (var change in parsed.Changes)
        {
            if (!byId.TryGetValue(change.Id, out var entry)) continue;

            // 바뀐 게 없으면 굳이 PowerPoint를 건드리지 않는다.
            if (!change.Insert && change.NewText == entry.Text) continue;

            changes.Add(new ParagraphChange
            {
                Id             = change.Id,
                SlideIndex     = entry.SlideIndex,
                ShapeName      = entry.ShapeName,
                ParagraphIndex = entry.ParagraphIndex,
                OldText        = change.Insert ? "" : entry.Text,
                NewText        = change.NewText,
                Kind           = change.Insert           ? ParagraphChangeKind.Insert
                               : change.NewText.Length == 0 ? ParagraphChangeKind.Clear
                               : ParagraphChangeKind.Replace,
            });
        }

        if (changes.Count == 0)
        {
            var message = parsed.Warnings.Count > 0 || parsed.UnmatchedLines.Count > 0
                ? "반영할 변경이 없습니다.\n\n" + string.Join("\n", parsed.Warnings.Take(5))
                : "변경된 항목이 없습니다.";
            MessageBox.Show(message, "텍스트 치환",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        AnnotateOverflow(changes);

        var preview = new ChangePreviewWindow(changes, parsed.Warnings, parsed.UnmatchedLines)
        {
            Owner = this
        };
        if (preview.ShowDialog() != true || preview.Accepted is not { Count: > 0 }) return;

        AppliedChanges = preview.Accepted;
        MarkRed        = preview.MarkRed;
        DialogResult   = true;
        Close();
    }

    /// <summary>
    /// 도형 단위로 변경 후 문단 구성을 만들어 넘침을 예측한다.
    /// 한 도형에 여러 문단이 바뀔 수 있으므로 개별 문단이 아니라 도형 전체로 봐야 한다.
    /// </summary>
    private void AnnotateOverflow(List<ParagraphChange> changes)
    {
        foreach (var group in changes.GroupBy(c => (c.SlideIndex, c.ShapeName)))
        {
            if (!_shapeLookup.TryGetValue(group.Key, out var shape)) continue;

            var paragraphs = _entries
                .Where(x => x.SlideIndex == group.Key.SlideIndex && x.ShapeName == group.Key.ShapeName)
                .OrderBy(x => x.ParagraphIndex)
                .Select(x => x.Text)
                .ToList();

            foreach (var change in group.Where(c => c.Kind != ParagraphChangeKind.Insert))
            {
                int i = change.ParagraphIndex - 1;
                if (i >= 0 && i < paragraphs.Count) paragraphs[i] = change.NewText;
            }
            foreach (var change in group.Where(c => c.Kind == ParagraphChangeKind.Insert)
                                        .OrderByDescending(c => c.ParagraphIndex))
            {
                int at = Math.Min(change.ParagraphIndex, paragraphs.Count);
                paragraphs.Insert(at, change.NewText);
            }

            var warning = OverflowEstimator.Estimate(shape, paragraphs);
            if (warning.Length == 0) continue;
            foreach (var change in group) change.OverflowWarning = warning;
        }
    }

    // ── 보조 ─────────────────────────────────────────────────────

    private void Reset_Click(object sender, RoutedEventArgs e)
        => Editor.Text = _originalBody;

    private void CopyAll_Click(object sender, RoutedEventArgs e)
        => Clipboard.SetText(Prompt + "\n\n" + _originalBody);

    private void CopyTextOnly_Click(object sender, RoutedEventArgs e)
        => Clipboard.SetText(Editor.Text);

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
}
