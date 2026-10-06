using System.Globalization;
using System.Text.RegularExpressions;

namespace BckPPT.Schedule;

public record ScheduleTask(string Name, DateTime Start, DateTime End, string StartText, string EndText);

public record ScheduleSection(string Name, IReadOnlyList<ScheduleTask> Tasks);

/// <summary>구분(단계). 하위 구분이 없으면 이름이 빈 섹션 하나를 가진다.</summary>
public record ScheduleGroup(string Number, string Name, IReadOnlyList<ScheduleSection> Sections)
{
    public IEnumerable<ScheduleTask> Tasks => Sections.SelectMany(s => s.Tasks);
    public bool HasSubsections => Sections.Any(s => s.Name.Length > 0);
}

public record ScheduleMilestone(string Name, DateTime Date, string DateText);

/// <summary>마일스톤 묶음. "## 운영: 보고회"면 Label=운영, Name=보고회.</summary>
public record MilestoneGroup(string Label, string Name, IReadOnlyList<ScheduleMilestone> Milestones);

public record ScheduleDocument(string Title, IReadOnlyList<ScheduleGroup> Groups, IReadOnlyList<MilestoneGroup> MilestoneGroups)
{
    public IEnumerable<ScheduleTask> AllTasks => Groups.SelectMany(g => g.Tasks);

    public DateTime First => AllTasks.Select(t => t.Start)
        .Concat(MilestoneGroups.SelectMany(m => m.Milestones).Select(m => m.Date)).Min();

    public DateTime Last => AllTasks.Select(t => t.End)
        .Concat(MilestoneGroups.SelectMany(m => m.Milestones).Select(m => m.Date)).Max();
}

/// <summary>
/// 일정 md를 읽는다. 형식(착수보고 43쪽에서 뽑은 형식):
/// <code>
/// # 제목
/// ## 1. 구분            ← 번호가 있으면 구분
/// ### 3-1. 하위 구분     ← 선택
/// | 항목 | 시작 | 끝 |    ← 세 칸 표는 일정
/// ## 운영: 보고회        ← 번호 없는 제목 아래 두 칸 표는 마일스톤
/// | 보고 | 시점 |
/// </code>
/// 날짜는 "2026-10-16", "2026.10.16", "2026-12 하순", "2026-12"(시작은 1일, 끝은 말일)을 받는다.
/// </summary>
public static class ScheduleMarkdown
{
    public static ScheduleDocument Parse(string markdown)
    {
        string title = "";
        var groups = new List<(string Number, string Name, List<(string Name, List<ScheduleTask> Tasks)> Sections)>();
        var milestoneGroups = new List<(string Label, string Name, List<ScheduleMilestone> Items)>();
        (string Number, string Name, List<(string Name, List<ScheduleTask> Tasks)> Sections)? group = null;
        (string Label, string Name, List<ScheduleMilestone> Items)? milestones = null;
        string section = "";

        foreach (var raw in markdown.Split('\n'))
        {
            var line = raw.Trim();
            if (line.StartsWith("# ")) { title = line[2..].Trim(); continue; }
            if (Regex.Match(line, @"^##\s+(\d+)[.)]\s*(.+)$") is { Success: true } g)
            {
                group = (g.Groups[1].Value, g.Groups[2].Value.Trim(), new());
                groups.Add(group.Value);
                milestones = null; section = "";
                continue;
            }
            if (line.StartsWith("## "))
            {
                var heading = line[3..].Trim();
                var parts = heading.Split(':', 2);
                milestones = parts.Length == 2 ? (parts[0].Trim(), parts[1].Trim(), new()) : (heading, heading, new());
                milestoneGroups.Add(milestones.Value);
                group = null;
                continue;
            }
            if (Regex.Match(line, @"^###\s+(?:[\d-]+[.)]\s*)?(.+)$") is { Success: true } s) { section = s.Groups[1].Value.Trim(); continue; }
            if (!line.StartsWith('|') || Regex.IsMatch(line, @"^\|[\s:|-]+\|?$")) continue;

            var cells = line.Trim('|').Split('|').Select(c => c.Trim()).ToArray();
            if (group is { } current && cells.Length >= 3
                && TryParseDate(cells[1], end: false, out var start) && TryParseDate(cells[2], end: true, out var finish))
            {
                var bucket = current.Sections.FirstOrDefault(x => x.Name == section);
                if (bucket.Tasks is null) { bucket = (section, new()); current.Sections.Add(bucket); }
                bucket.Tasks.Add(new ScheduleTask(cells[0], start, finish, cells[1], cells[2]));
            }
            else if (milestones is { } m && cells.Length >= 2 && TryParseDate(cells[1], end: false, out var at, milestone: true))
                m.Items.Add(new ScheduleMilestone(cells[0], at, cells[1]));
        }

        return new ScheduleDocument(title,
            groups.Where(x => x.Sections.Count > 0)
                .Select(x => new ScheduleGroup(x.Number, x.Name, x.Sections.Select(s => new ScheduleSection(s.Name, s.Tasks)).ToList())).ToList(),
            milestoneGroups.Where(x => x.Items.Count > 0).Select(x => new MilestoneGroup(x.Label, x.Name, x.Items)).ToList());
    }

    /// <summary>
    /// 날짜 한 칸을 읽는다. 상순·중순·하순은 시작이면 1·11·21일, 끝이면 10·20·말일. 마일스톤이면 5·15·25일(가운데).
    /// </summary>
    public static bool TryParseDate(string text, bool end, out DateTime date, bool milestone = false)
    {
        date = default;
        text = text.Trim();
        if (Regex.Match(text, @"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})\.?$") is { Success: true } d)
            return TryMake(d, int.Parse(d.Groups[3].Value), out date);
        if (Regex.Match(text, @"^(\d{4})[-./](\d{1,2})\.?\s*(상순|초|중순|중|하순|말)?$") is { Success: true } m)
        {
            int year = int.Parse(m.Groups[1].Value), month = int.Parse(m.Groups[2].Value);
            if (month is < 1 or > 12) return false;
            int last = DateTime.DaysInMonth(year, month);
            int day = m.Groups[3].Value switch
            {
                "상순" or "초" => milestone ? 5 : end ? 10 : 1,
                "중순" or "중" => milestone ? 15 : end ? 20 : 11,
                "하순" or "말" => milestone ? 25 : end ? last : 21,
                _ => milestone ? 15 : end ? last : 1,
            };
            date = new DateTime(year, month, day);
            return true;
        }
        return false;

        static bool TryMake(Match match, int day, out DateTime value) =>
            DateTime.TryParseExact($"{match.Groups[1].Value}-{int.Parse(match.Groups[2].Value):00}-{day:00}", "yyyy-MM-dd",
                CultureInfo.InvariantCulture, DateTimeStyles.None, out value);
    }
}
