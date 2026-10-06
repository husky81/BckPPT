using BckPPT.Models;
using BckPPT.Services;

namespace BckPPT.Schedule;

public enum ScheduleStyle
{
    /// <summary>단계 | 세부 추진 항목 | 기간 | 월 눈금(막대). 항목 이름과 기간을 표 칸에 적는다.</summary>
    Table,
    /// <summary>단계 | 월 눈금. 항목 이름을 막대 옆에 적는다(착수보고 43쪽 형태). 눈금 칸이 넓다.</summary>
    Gantt,
    /// <summary>단계 | 월 눈금. 항목 이름을 두꺼운 막대 안에 적는다. 막대가 짧아 안 들어가면 옆에 적는다.</summary>
    InBar,
}

/// <summary>그릴 영역(pt). 없으면 16:9(960×540) 슬라이드의 제목 아래 본문 영역.</summary>
public record ScheduleArea(float Left, float Top, float Width, float Height)
{
    public static ScheduleArea Default { get; } = new(33, 80, 897, 424);
}

public record ScheduleOptions(
    ScheduleStyle Style      = ScheduleStyle.Table,
    ScheduleArea? Area       = null,
    IReadOnlyList<string>? Palette = null,
    string        HeaderFill = "#0A3671",
    string        SubHeaderFill = "#DCE4EF",
    string        BandFill   = "#F3F6FB",
    string        GridColor  = "#D3DCEA",
    string        Accent     = "#E46C0A",
    string        TextColor  = "#1F2937",
    string        FontName   = "");

public record ScheduleResult(string Name, string TableName, int Rows, int Periods, int OverlayShapes, string Summary);

/// <summary>
/// 일정 md(<see cref="ScheduleDocument"/>)로 일정표를 그린다. 눈금·구분·항목은 PowerPoint 표로 만들어 사람이 고치기 쉽게 하고,
/// 막대·마일스톤은 표의 실제 행·열 위치에 맞춰 도형으로 얹은 뒤 표와 함께 그룹으로 묶는다.
/// 배치는 모두 코드로 계산한다(LLM 없음).
/// </summary>
public sealed class ScheduleRenderer(IPptComService ppt, string presentationPath)
{
    public static readonly IReadOnlyList<string> DefaultPalette = ["#0A3671", "#1F4E8C", "#2F6BB3", "#4A86C8", "#5B9BD5", "#7FA7D9"];
    private const string MilestoneLabelFill = "#5B6B85";

    public record Period(string Label, int Year, DateTime Start, DateTime End);   // End는 다음 칸 시작(배타)

    private enum RowKind { Task, Milestones }
    private sealed record BodyRow(RowKind Kind, int Group, int Section, ScheduleTask? Task, MilestoneGroup? Milestones);

    public ScheduleResult Draw(int slide, ScheduleDocument doc, ScheduleOptions? options = null)
    {
        options ??= new ScheduleOptions();
        if (!doc.AllTasks.Any() && doc.MilestoneGroups.Count == 0) throw new ArgumentException("일정 항목이 없습니다.");
        var area = options.Area ?? ScheduleArea.Default;
        var palette = options.Palette is { Count: > 0 } p ? p : DefaultPalette;
        var periods = Periods(doc.First, doc.Last);
        string prefix = $"일정{DateTime.Now:HHmmss}";

        var rows = new List<BodyRow>();
        for (int g = 0; g < doc.Groups.Count; g++)
            for (int s = 0; s < doc.Groups[g].Sections.Count; s++)
                foreach (var task in doc.Groups[g].Sections[s].Tasks) rows.Add(new BodyRow(RowKind.Task, g, s, task, null));
        foreach (var m in doc.MilestoneGroups) rows.Add(new BodyRow(RowKind.Milestones, -1, 0, null, m));

        bool table = options.Style == ScheduleStyle.Table;
        bool inBar = options.Style == ScheduleStyle.InBar;
        bool hasSub = doc.Groups.Any(g => g.HasSubsections);

        // ── 열 ──
        var widths = new List<float>();
        float groupTotal = area.Width * (table ? 0.10f : 0.12f);
        int colGroup = 0, colSub = -1, colItem = -1, colPeriodText = -1;
        if (hasSub) { widths.Add(groupTotal * 0.56f); widths.Add(groupTotal * 0.44f); colSub = 1; }
        else widths.Add(groupTotal);
        if (table)
        {
            colItem = widths.Count; widths.Add(area.Width * 0.245f);
            colPeriodText = widths.Count; widths.Add(area.Width * 0.07f);
        }
        int firstPeriodCol = widths.Count;
        float periodW = (area.Width - widths.Sum()) / periods.Count;
        widths.AddRange(Enumerable.Repeat(periodW, periods.Count));
        int cols = widths.Count;

        // ── 행 높이·글자 ──
        float headH = Math.Min(18f, area.Height / (rows.Count + 2));
        float bodyH = (area.Height - headH * 2) / rows.Count;
        float font = Math.Clamp(bodyH * 0.44f, 7f, 10.5f);

        // ── 칸 ──
        var cells = Enumerable.Range(0, rows.Count + 2).Select(_ => Enumerable.Repeat("", cols).ToArray()).ToList();
        var merges = new List<TableMerge>();
        var styles = new List<TableCellStyle>();
        int groupLastCol = hasSub ? colSub : colGroup;

        cells[0][colGroup] = "단계";
        merges.Add(new TableMerge(1, colGroup + 1, 2, groupLastCol + 1));
        if (table)
        {
            cells[0][colItem] = "세부 추진 항목"; merges.Add(new TableMerge(1, colItem + 1, 2, colItem + 1));
            cells[0][colPeriodText] = "기간"; merges.Add(new TableMerge(1, colPeriodText + 1, 2, colPeriodText + 1));
        }
        for (int i = 0; i < periods.Count;)
        {
            int span = 0;
            while (i + span < periods.Count && periods[i + span].Year == periods[i].Year) span++;
            cells[0][firstPeriodCol + i] = periods[i].Year.ToString();
            if (span > 1) merges.Add(new TableMerge(1, firstPeriodCol + i + 1, 1, firstPeriodCol + i + span));
            i += span;
        }
        for (int i = 0; i < periods.Count; i++)
        {
            cells[1][firstPeriodCol + i] = periods[i].Label;
            styles.Add(new TableCellStyle(2, firstPeriodCol + i + 1, Fill: options.SubHeaderFill, FontColor: options.TextColor, Bold: true));
        }

        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            int tr = r + 3;   // 표의 행 번호(1부터, 머리 2행)
            if (row.Kind == RowKind.Task)
            {
                if (table)
                {
                    cells[r + 2][colItem] = row.Task!.Name;
                    cells[r + 2][colPeriodText] = $"{row.Task.Start:M.d}~{row.Task.End:M.d}";
                }
            }
            else
            {
                var m = row.Milestones!;
                cells[r + 2][colGroup] = m.Label;
                if (hasSub) merges.Add(new TableMerge(tr, colGroup + 1, tr, colSub + 1));
                styles.Add(new TableCellStyle(tr, colGroup + 1, Fill: MilestoneLabelFill, FontColor: "#FFFFFF", Bold: true));
                if (table)
                {
                    cells[r + 2][colItem] = m.Name;
                    cells[r + 2][colPeriodText] = $"{m.Milestones.Count}회";
                }
            }
        }
        // 구분·하위 구분 칸 병합
        for (int g = 0; g < doc.Groups.Count; g++)
        {
            int first = rows.FindIndex(x => x.Kind == RowKind.Task && x.Group == g), last = rows.FindLastIndex(x => x.Kind == RowKind.Task && x.Group == g);
            var group = doc.Groups[g];
            string color = palette[g % palette.Count];
            // 좁은 칸에서 낱말 중간이 잘리지 않게(예: "3. 단기 상 / 세기획") 띄어쓰기에서 균형 있게 나눈다.
            string label = $"{group.Number}. {group.Name}";
            float labelWidth = widths[colGroup] + (!group.HasSubsections && hasSub ? widths[colSub] : 0) - 8f;
            if (Estimate(label, font) > labelWidth) label = Agent.Cards.CardFillPrompts.BalancedBreak(label) ?? label;
            cells[first + 2][colGroup] = label;
            bool split = group.HasSubsections;
            merges.Add(new TableMerge(first + 3, colGroup + 1, last + 3, split || !hasSub ? colGroup + 1 : colSub + 1));
            styles.Add(new TableCellStyle(first + 3, colGroup + 1, Fill: color, FontColor: "#FFFFFF", Bold: true));
            if (!split) continue;
            for (int s = 0; s < group.Sections.Count; s++)
            {
                int sf = rows.FindIndex(x => x.Kind == RowKind.Task && x.Group == g && x.Section == s);
                int sl = rows.FindLastIndex(x => x.Kind == RowKind.Task && x.Group == g && x.Section == s);
                cells[sf + 2][colSub] = group.Sections[s].Name;
                merges.Add(new TableMerge(sf + 3, colSub + 1, sl + 3, colSub + 1));
                styles.Add(new TableCellStyle(sf + 3, colSub + 1, Fill: Lighten(color, 0.35f), FontColor: "#FFFFFF", Bold: true,
                    Align: TextAlign.Center, FontSize: font - 0.5f));
            }
        }

        var aligns = Enumerable.Range(0, cols).Select(c =>
            c == colGroup ? TextAlign.Left : c == colSub ? TextAlign.Center : c == colItem ? TextAlign.Left : TextAlign.Center).ToList();
        var layout = ppt.AddTable(presentationPath, slide, new TableRequest(
            area.Left, area.Top, area.Width, cells.Select(r => (IReadOnlyList<string>)r).ToList(),
            ColumnWidths: widths, FontSize: font, FontName: options.FontName, HeaderRows: 2,
            HeaderFill: options.HeaderFill, HeaderFontColor: "#FFFFFF", FontColor: options.TextColor,
            BandFill: options.BandFill, BorderColor: options.GridColor, BorderWidth: 0.5f,
            ColumnAligns: aligns, Merges: merges, CellStyles: styles,
            RowHeights: [headH, headH, .. Enumerable.Repeat(bodyH, rows.Count)], CellPadding: 4f, Name: prefix + "_표"));

        // ── 막대·마일스톤 ──
        float gridLeft = layout.ColumnLefts[firstPeriodCol];
        float gridRight = layout.ColumnLefts[^1] + layout.ColumnWidths[^1];
        float X(DateTime date)
        {
            for (int i = 0; i < periods.Count; i++)
            {
                if (date >= periods[i].End && i < periods.Count - 1) continue;
                double frac = (date - periods[i].Start).TotalDays / (periods[i].End - periods[i].Start).TotalDays;
                return layout.ColumnLefts[firstPeriodCol + i] + (float)Math.Clamp(frac, 0, 1) * layout.ColumnWidths[firstPeriodCol + i];
            }
            return gridRight;
        }

        var overlay = new List<TextShapeRequest>();
        int n = 0;
        string Next() => $"{prefix}_{++n}";
        for (int r = 0; r < rows.Count; r++)
        {
            var row = rows[r];
            float top = layout.RowTops[r + 2], h = layout.RowHeights[r + 2], mid = top + h / 2;
            if (row.Kind == RowKind.Task)
            {
                var task = row.Task!;
                float x1 = X(task.Start), x2 = Math.Max(X(task.End.AddDays(1)), x1 + 3f);
                string color = palette[row.Group % palette.Count];
                if (inBar)
                {
                    // 이름을 막대 안에 쓴다. 안 들어가면 글자를 6.5pt까지 줄이고, 그래도 안 되면 막대 옆에 쓴다.
                    float barH = h * 0.78f, barW = x2 - x1;
                    // 어림은 실제 글꼴보다 좁게 나와 두 줄로 넘친 막대가 있었다(2026-10-07). 10% 여유를 둔다.
                    bool Fits(float sz) => Estimate(task.Name, sz) * 1.1f + 8f <= barW;
                    float size = font;
                    while (size > 6.5f && !Fits(size)) size -= 0.5f;
                    if (Fits(size))
                    {
                        overlay.Add(new TextShapeRequest(task.Name, x1, mid - barH / 2, barW, barH, Next(), TextShapeKind.RoundedRectangle, color,
                            FontName: options.FontName, FontSize: size, FontColor: "#FFFFFF", Padding: 3f, Align: TextAlign.Center));
                        continue;
                    }
                    overlay.Add(new TextShapeRequest("", x1, mid - barH / 2, barW, barH, Next(), TextShapeKind.RoundedRectangle, color));
                }
                else
                {
                    float barH = Math.Min(table ? 9f : 8f, h * 0.45f);
                    overlay.Add(new TextShapeRequest("", x1, mid - barH / 2, x2 - x1, barH, Next(), TextShapeKind.RoundedRectangle, color));
                }
                if (!table)
                {
                    // 이름은 막대 오른쪽, 자리가 없으면 왼쪽.
                    float est = Estimate(task.Name, font) + 4f;
                    bool right = x2 + 3 + est <= gridRight || x1 - 3 - est < gridLeft;
                    overlay.Add(new TextShapeRequest(task.Name, right ? x2 + 3 : x1 - 3 - est, top, est, h, Next(),
                        FontName: options.FontName, FontSize: font, FontColor: options.TextColor, AutoWidth: true,
                        RightEdge: right ? null : x1 - 3, Padding: 0, Align: right ? TextAlign.Left : TextAlign.Right));
                }
            }
            else
            {
                float d = Math.Min(10f, h * 0.6f);
                foreach (var m in row.Milestones!.Milestones)
                {
                    float x = X(m.Date);
                    overlay.Add(new TextShapeRequest("", x - d / 2, mid - d / 2, d, d, Next(), TextShapeKind.Diamond, options.Accent));
                    float est = Estimate(m.Name, font) + 2f;
                    bool right = x + d / 2 + 2 + est <= gridRight;
                    overlay.Add(new TextShapeRequest(m.Name, right ? x + d / 2 + 2 : x - d / 2 - 2 - est, top, est, h, Next(),
                        FontName: options.FontName, FontSize: font, Bold: true, FontColor: options.Accent, AutoWidth: true,
                        RightEdge: right ? null : x - d / 2 - 2, Padding: 0, Align: right ? TextAlign.Left : TextAlign.Right));
                }
            }
        }
        var names = ppt.AddTextShapes(presentationPath, slide, overlay);

        // 표와 얹은 도형을 한 그룹으로. 표를 묶을 수 없는 버전이면 얹은 도형만 묶는다.
        string groupName;
        try { groupName = ppt.GroupShapes(presentationPath, slide, [layout.Name, .. names], prefix); }
        catch (Exception) when (names.Count >= 2) { groupName = ppt.GroupShapes(presentationPath, slide, names, prefix + "_막대"); }

        var firstStart = periods[0].Start; var lastP = periods[^1];
        string unit = periods[0].Label.EndsWith('Q') ? "분기" : "월";
        string summary = $"일정표({(table ? "표형" : inBar ? "막대 안 이름형" : "간트형")}): 구분 {doc.Groups.Count}개·항목 {doc.AllTasks.Count()}개·마일스톤 {doc.MilestoneGroups.Sum(m => m.Milestones.Count)}개, " +
                         $"{firstStart:yyyy-MM}~{lastP.End.AddDays(-1):yyyy-MM}({unit} {periods.Count}칸). 표 '{layout.Name}', 막대·표시 {names.Count}개, 그룹 '{groupName}'";
        return new ScheduleResult(groupName, layout.Name, rows.Count, periods.Count, names.Count, summary);
    }

    /// <summary>눈금 칸. 18개월 이하면 월, 넘으면 분기.</summary>
    public static IReadOnlyList<Period> Periods(DateTime first, DateTime last)
    {
        var start = new DateTime(first.Year, first.Month, 1);
        int months = (last.Year - start.Year) * 12 + last.Month - start.Month + 1;
        var result = new List<Period>();
        if (months <= 18)
        {
            for (int i = 0; i < months; i++)
            {
                var s = start.AddMonths(i);
                result.Add(new Period($"{s.Month}월", s.Year, s, s.AddMonths(1)));
            }
            return result;
        }
        var q = new DateTime(start.Year, (start.Month - 1) / 3 * 3 + 1, 1);
        for (; q <= last; q = q.AddMonths(3))
            result.Add(new Period($"{(q.Month - 1) / 3 + 1}Q", q.Year, q, q.AddMonths(3)));
        return result;
    }

    /// <summary>글자 폭 어림(pt). 한글은 글자 크기, 영숫자는 그 절반 남짓.</summary>
    public static float Estimate(string text, float size) =>
        text.Split('\v', '\n').Max(line => line.Sum(c =>
            c >= 0x1100 ? size * 0.95f : c == ' ' ? size * 0.3f : char.IsUpper(c) || char.IsDigit(c) ? size * 0.6f : size * 0.5f));

    public static string Lighten(string hex, float amount)
    {
        int v = Convert.ToInt32(hex.TrimStart('#'), 16);
        int Mix(int c) => (int)Math.Round(c + (255 - c) * amount);
        return $"#{Mix(v >> 16 & 255):X2}{Mix(v >> 8 & 255):X2}{Mix(v & 255):X2}";
    }
}
