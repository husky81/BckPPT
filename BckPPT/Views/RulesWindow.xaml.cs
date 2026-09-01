using System.Collections.ObjectModel;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using CellBox2.BckPPT.Models;
using CellBox2.BckPPT.Services;

namespace CellBox2.BckPPT.Views;

public partial class RulesWindow : Window
{
    private readonly ObservableCollection<ReplaceRule> _rules;

    public RulesWindow()
    {
        InitializeComponent();

        SourceInitialized += (_, _) =>
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            int value = 1;
            DwmSetWindowAttribute(hwnd, 20, ref value, sizeof(int));
        };

        // 기존 규칙 로드
        _rules = new ObservableCollection<ReplaceRule>(
            PptRules.Current.ReplaceRules.Select(r => new ReplaceRule
            {
                Find          = r.Find,
                Replace       = r.Replace,
                CaseSensitive = r.CaseSensitive,
            }));

        RulesGrid.ItemsSource = _rules;
    }

    private void AddRow_Click(object sender, RoutedEventArgs e)
    {
        _rules.Add(new ReplaceRule());
        RulesGrid.ScrollIntoView(_rules.Last());
        RulesGrid.SelectedItem = _rules.Last();
    }

    private void Save_Click(object sender, RoutedEventArgs e)
    {
        // DataGrid에서 편집 커밋
        RulesGrid.CommitEdit(System.Windows.Controls.DataGridEditingUnit.Row, true);

        var rules = PptRules.Current;
        rules.ReplaceRules = _rules
            .Where(r => !string.IsNullOrWhiteSpace(r.Find))
            .ToList();
        rules.Save();

        MessageBox.Show($"{rules.ReplaceRules.Count}개 규칙이 저장되었습니다.",
            "저장 완료", MessageBoxButton.OK, MessageBoxImage.Information);
        DialogResult = true;
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
}
