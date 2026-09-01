using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Interop;

namespace CellBox2.BckPPT.Views;

public partial class TextEditorWindow : Window
{
    private readonly List<(int Seq, int SlideIndex, string ShapeName, string FullText)> _entries;
    private readonly string _presentationPath;
    private readonly StringBuilder _originalBody = new();

    /// <summary>적용된 변경 목록 (OldText, NewText는 실제 줄바꿈 \r 포함). null이면 취소됨.</summary>
    public List<(int SlideIndex, string ShapeName, string OldText, string NewText)>? AppliedChanges
        { get; private set; }

    public TextEditorWindow(
        string presentationPath,
        List<(int Seq, int SlideIndex, string ShapeName, string FullText)> entries,
        string slideTitle)
    {
        InitializeComponent();
        _entries          = entries;
        _presentationPath = presentationPath;
        Title             = $"텍스트 치환 — {slideTitle}";

        SourceInitialized += (_, _) =>
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            int value = 1;
            DwmSetWindowAttribute(hwnd, 20, ref value, sizeof(int));
        };

        PromptBox.Text =
            "# 아래 PPT 슬라이드 텍스트를 사용자 요청에 맞게 수정해줘.\n" +
            "- 텍스트 추가/삭제 불가, 기존 내용 수정만 가능\n" +
            "- 수정이 필요한 항목만 \"번호: 수정내용\" 형식으로 작성\n" +
            "- 번호는 반드시 원본과 동일하게 유지\n" +
            "- 줄바꿈은 \\n으로 표시";

        // 번호:텍스트 목록 생성 (\r 계열을 모두 \\n으로 이스케이프)
        foreach (var (seq, _, _, text) in entries)
        {
            var escaped = EscapeNewlines(text);
            _originalBody.AppendLine($"{seq}: {escaped}");
        }
        Editor.Text = _originalBody.ToString();
    }

    // ── \r\n / \r / \n → \\n (표시용) ──────────────────────────

    private static string EscapeNewlines(string s)
        => s.Replace("\r\n", "\\n").Replace("\r", "\\n").Replace("\n", "\\n");

    /// <summary>\\n → \r (PPT 내부 줄바꿈).</summary>
    private static string UnescapeNewlines(string s)
        => s.Replace("\\n", "\r");

    // ── 적용 ─────────────────────────────────────────────────────

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        var changes = new List<(int SlideIndex, string ShapeName, string OldText, string NewText)>();
        var lines   = Editor.Text.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        foreach (var line in lines)
        {
            var trimmed = line.TrimEnd('\r');           // \r만 제거, 공백은 보존
            var match   = Regex.Match(trimmed, @"^(\d+):\s*(.*)$");
            if (!match.Success) continue;

            int    seq           = int.Parse(match.Groups[1].Value);
            string editedEscaped = match.Groups[2].Value;

            var entry = _entries.FirstOrDefault(x => x.Seq == seq);
            if (entry.ShapeName is null) continue;

            // 원본 이스케이프와 비교 — 동일하면 스킵
            var originalEscaped = EscapeNewlines(entry.FullText);
            if (editedEscaped == originalEscaped) continue;

            // 실제 적용할 텍스트 (\\n → \r)
            string newText = UnescapeNewlines(editedEscaped);
            changes.Add((entry.SlideIndex, entry.ShapeName, entry.FullText, newText));
        }

        if (changes.Count == 0)
        {
            MessageBox.Show("변경된 항목이 없습니다.", "텍스트 치환",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        // Diff 확인 (표시용은 이스케이프된 형태)
        var diffSb = new StringBuilder();
        diffSb.AppendLine($"{changes.Count}건 변경 확인:\n");
        foreach (var (_, shapeName, oldText, newText) in changes)
        {
            diffSb.AppendLine($"[{shapeName}]");
            diffSb.AppendLine($"  - {EscapeNewlines(oldText)}");
            diffSb.AppendLine($"  + {EscapeNewlines(newText)}");
            diffSb.AppendLine();
        }
        diffSb.Append("적용하시겠습니까?");

        if (MessageBox.Show(diffSb.ToString(), "변경 내용 확인",
                MessageBoxButton.YesNo, MessageBoxImage.Question)
            != MessageBoxResult.Yes) return;

        // 변경 목록만 전달 — 실제 적용은 ViewModel(ApplyTextChangesAsync)에서 담당
        AppliedChanges = changes;
        DialogResult   = true;
        Close();
    }

    private void CopyAll_Click(object sender, RoutedEventArgs e)
        => Clipboard.SetText(PromptBox.Text + "\n\n" + _originalBody.ToString());

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
