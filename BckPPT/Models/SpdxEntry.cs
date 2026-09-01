namespace CellBox2.BckPPT.Models;

/// <summary>SPDX(Slide-MPDX) 노드 하나. TSV 직렬화 + apply-back 메타.</summary>
public record SpdxEntry(
    int Id,
    int ParentId,
    string Type,        // "slide", "shape:제목", "line", "lineB"
    int MaxHW,          // line 노드: 최대 반각 글자 수, 그 외 0
    float FontSize,     // line 노드: 글꼴 크기 (pt), 그 외 0
    string RoleHint,    // line 노드: "title" / "subtitle" / "body", 그 외 ""
    string Text,
    // ── apply-back 메타 (TSV에 미포함) ──
    int SlideIndex,
    string ShapeName,
    int? LineIndex);    // line 노드일 때 0-based 줄 인덱스
