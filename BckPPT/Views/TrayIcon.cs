using System.Drawing;
using System.Windows.Forms;

namespace BckPPT.Views;

/// <summary>
/// 작업 표시줄 알림 영역 아이콘. BckPPT는 창 없이 여기서 시작하고, MCP 서버가 뒤에서 외부 에이전트를 받는다.
/// 두 번 누르면 채팅창, 오른쪽 메뉴에서 편집창·MCP 연결 명령 복사·종료.
/// </summary>
public sealed class TrayIcon : IDisposable
{
    private readonly NotifyIcon _icon;

    public TrayIcon(string mcpStatus, string mcpCommand, Action openChat, Action openEditor, Action exit)
    {
        var menu = new ContextMenuStrip();
        var chat = new ToolStripMenuItem("채팅창 열기", null, (_, _) => openChat()) { Font = new Font(menu.Font, FontStyle.Bold) };
        menu.Items.Add(chat);
        menu.Items.Add(new ToolStripMenuItem("전문 편집창", null, (_, _) => openEditor()));
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(new ToolStripMenuItem(mcpStatus) { Enabled = false });
        menu.Items.Add(new ToolStripMenuItem("Claude Code 연결 명령 복사", null, (_, _) =>
        {
            try
            {
                Clipboard.SetText(mcpCommand);
                ShowBalloon("복사했습니다", mcpCommand);
            }
            catch { }
        }));
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add(new ToolStripMenuItem("종료", null, (_, _) => exit()));

        _icon = new NotifyIcon
        {
            Icon = LoadIcon(),
            Text = "BckPPT",
            ContextMenuStrip = menu,
            Visible = true,
        };
        _icon.DoubleClick += (_, _) => openChat();
    }

    public void ShowBalloon(string title, string text) =>
        _icon.ShowBalloonTip(4000, title, text, ToolTipIcon.None);

    private static Icon LoadIcon()
    {
        try
        {
            var resource = System.Windows.Application.GetResourceStream(new Uri("pack://application:,,,/Assets/icon.ico"));
            if (resource is not null) return new Icon(resource.Stream);
        }
        catch { }
        return SystemIcons.Application;
    }

    public void Dispose()
    {
        _icon.Visible = false;
        _icon.Dispose();
    }
}
