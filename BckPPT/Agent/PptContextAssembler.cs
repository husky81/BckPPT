using System.Text;
using BckPPT.Application;
using BckPPT.Models;

namespace BckPPT.Agent;

public static class PptContextAssembler
{
    private const int MaxShapes = 120;
    private const int MaxTextPerShape = 800;
    private const int MaxContextCharacters = 40_000;

    public static string Assemble(DocumentSession session, PresentationAnalysis analysis)
    {
        var builder = new StringBuilder();
        builder.AppendLine($"파일명: {session.Presentation?.Name}");
        builder.AppendLine($"슬라이드 수: {analysis.Slides.Count(item => !item.IsAll)}");
        if (session.SelectedSlideIndex is not null) builder.AppendLine($"선택 슬라이드: {session.SelectedSlideIndex}");
        builder.AppendLine("Shape 구조:");

        var shapes = analysis.Shapes
            .OrderBy(shape => session.SelectedSlideIndex == shape.SlideIndex ? 0 : 1)
            .ThenBy(shape => shape.SlideIndex)
            .Take(MaxShapes);
        foreach (var shape in shapes)
        {
            var text = shape.FullText.Replace('\v', '\n');
            if (text.Length > MaxTextPerShape) text = text[..MaxTextPerShape] + "…";
            builder.Append("- slide=").Append(shape.SlideIndex)
                .Append("; shape=").Append(shape.ShapeName)
                .Append("; role=").Append(shape.Role)
                .Append("; lines=").Append(shape.LineCount)
                .Append("; overflow=").Append(shape.IsOverflow)
                .Append("; text=").AppendLine(text.Replace("\r", "\\r").Replace("\n", "\\n"));
            if (builder.Length >= MaxContextCharacters) break;
        }
        if (analysis.Shapes.Count > MaxShapes || builder.Length >= MaxContextCharacters)
            builder.AppendLine("(문맥 크기 제한으로 일부 Shape 생략)");
        return builder.ToString()[..Math.Min(builder.Length, MaxContextCharacters)];
    }
}
