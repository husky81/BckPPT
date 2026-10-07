using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;

namespace BckPPT.Views;

/// <summary>
/// 모델 답의 가벼운 마크다운(**굵게**, `코드`)을 TextBlock 글자 서식으로 바꾼다. 표·제목 같은 것은 글자 그대로 둔다.
/// qwen 답에 "**AI 기반 TBM**"처럼 별표가 그대로 보였다(2026-10-07).
/// </summary>
public static class MarkdownText
{
    public static readonly DependencyProperty TextProperty = DependencyProperty.RegisterAttached(
        "Text", typeof(string), typeof(MarkdownText), new PropertyMetadata("", OnTextChanged));

    public static string GetText(DependencyObject d) => (string)d.GetValue(TextProperty);
    public static void SetText(DependencyObject d, string value) => d.SetValue(TextProperty, value);

    private static readonly Regex Token = new(@"\*\*(.+?)\*\*|`([^`]+)`", RegexOptions.Compiled);

    private static void OnTextChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not TextBlock block) return;
        block.Inlines.Clear();
        var text = (e.NewValue as string ?? "").Replace("\r\n", "\n");
        // 줄 앞 "### " 같은 제목 표시는 떼고 굵게 쓴다.
        text = Regex.Replace(text, @"(?m)^#{1,6}\s+(.+)$", m => "**" + m.Groups[1].Value.Replace("**", "").Trim() + "**");
        int at = 0;
        foreach (Match m in Token.Matches(text))
        {
            if (m.Index > at) block.Inlines.Add(new Run(text[at..m.Index]));
            if (m.Groups[1].Success) block.Inlines.Add(new Bold(new Run(m.Groups[1].Value)));
            else block.Inlines.Add(new Run(m.Groups[2].Value) { FontFamily = (System.Windows.Media.FontFamily)block.FindResource("MonoFont") });
            at = m.Index + m.Length;
        }
        if (at < text.Length) block.Inlines.Add(new Run(text[at..]));
    }
}
