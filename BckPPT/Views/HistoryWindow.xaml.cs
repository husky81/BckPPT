using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Views;

// DataGrid 바인딩용 뷰모델
internal class HistoryRowVm(ChangeHistoryEntry e)
{
    public string TimestampStr      { get; } = e.Timestamp.ToString("HH:mm:ss");
    public string PresentationName  { get; } = e.PresentationName;
    public string ChangeType        { get; } = e.ChangeType;
    public int    SlideIndex        { get; } = e.SlideIndex;
    public string ShapeName         { get; } = e.ShapeName;
    public string OldValueShort     { get; } = Truncate(e.OldValue, 50);
    public string NewValueShort     { get; } = Truncate(e.NewValue, 50);

    private static string Truncate(string s, int max)
        => s.Length <= max ? s : s[..(max - 1)] + "…";
}

public partial class HistoryWindow : Window
{
    public HistoryWindow()
    {
        InitializeComponent();
        Title = "변경 이력";

        SourceInitialized += (_, _) =>
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            int value = 1;
            DwmSetWindowAttribute(hwnd, 20, ref value, sizeof(int));
        };

        var entries = ChangeHistoryService.LoadRecent(500);
        HistoryGrid.ItemsSource = entries
            .Select(e => new HistoryRowVm(e))
            .ToList();
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();

    private void OpenFolder_Click(object sender, RoutedEventArgs e)
    {
        var dir = ChangeHistoryService.GetHistoryDirectory();
        if (System.IO.Directory.Exists(dir))
            Process.Start(new ProcessStartInfo("explorer.exe", dir) { UseShellExecute = true });
    }

    [DllImport("dwmapi.dll", PreserveSig = true)]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);
}
