using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace BckPPT.Views;

public class DropGuideAdorner : Adorner
{
    public enum DropPosition { None, Above, Inside, Below }

    public DropPosition Position { get; set; }
    public Rect TargetRect { get; set; }

    private readonly Pen _linePen;
    private readonly Brush _highlightBrush;

    public DropGuideAdorner(UIElement adornedElement)
        : base(adornedElement)
    {
        IsHitTestVisible = false;
        _linePen = new Pen(new SolidColorBrush(Color.FromRgb(0, 122, 204)), 2);
        _highlightBrush = new SolidColorBrush(Color.FromArgb(50, 0, 122, 204));
    }

    protected override void OnRender(DrawingContext dc)
    {
        if (Position == DropPosition.None) return;

        switch (Position)
        {
            case DropPosition.Above:
                dc.DrawLine(_linePen,
                    new Point(TargetRect.Left, TargetRect.Top),
                    new Point(TargetRect.Right, TargetRect.Top));
                break;

            case DropPosition.Below:
                dc.DrawLine(_linePen,
                    new Point(TargetRect.Left, TargetRect.Bottom),
                    new Point(TargetRect.Right, TargetRect.Bottom));
                break;

            case DropPosition.Inside:
                dc.DrawRectangle(_highlightBrush, _linePen, TargetRect);
                break;
        }
    }
}
