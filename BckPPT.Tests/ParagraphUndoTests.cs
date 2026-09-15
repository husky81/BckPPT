using BckPPT.Models;
using BckPPT.ViewModels;

namespace BckPPT.Tests;

/// <summary>
/// 되돌리기는 "적용 후" 문서를 상대하므로 인덱스 보정이 핵심이다.
/// 실제로 이 보정이 빠져 문단이 한 칸씩 밀리는 버그가 있었다.
/// </summary>
public class ParagraphUndoTests
{
    private static ParagraphChange Change(
        int paragraphIndex, string oldText, string newText, ParagraphChangeKind kind) => new()
    {
        Id             = $"S1-1-{paragraphIndex}",
        SlideIndex     = 1,
        ShapeName      = "Pentagon 95",
        ParagraphIndex = paragraphIndex,
        OldText        = oldText,
        NewText        = newText,
        Kind           = kind,
    };

    [Fact]
    public void 빈문단_지시는_삭제로_다룬다()
    {
        Assert.Equal(ParagraphEditKind.Delete,
            BckPptViewModel.ToEditKind(ParagraphChangeKind.Clear));
        Assert.Equal(ParagraphEditKind.Insert,
            BckPptViewModel.ToEditKind(ParagraphChangeKind.Insert));
        Assert.Equal(ParagraphEditKind.Replace,
            BckPptViewModel.ToEditKind(ParagraphChangeKind.Replace));
    }

    [Fact]
    public void 단순교체는_같은_인덱스로_원문을_되돌린다()
    {
        var undo = BckPptViewModel.BuildUndoEdits(
        [
            Change(1, "옛1", "새1", ParagraphChangeKind.Replace),
            Change(3, "옛3", "새3", ParagraphChangeKind.Replace),
        ]);

        Assert.Equal(2, undo.Count);
        Assert.All(undo, e => Assert.Equal(ParagraphEditKind.Replace, e.Kind));
        Assert.Equal((1, "옛1"), (undo[0].ParagraphIndex, undo[0].NewText));
        Assert.Equal((3, "옛3"), (undo[1].ParagraphIndex, undo[1].NewText));
    }

    [Fact]
    public void 삽입은_생성된_문단을_지워서_되돌린다()
    {
        var undo = BckPptViewModel.BuildUndoEdits(
            [Change(1, "", "추가된 줄", ParagraphChangeKind.Insert)]);

        var edit = Assert.Single(undo);
        Assert.Equal(ParagraphEditKind.Delete, edit.Kind);
        Assert.Equal(2, edit.ParagraphIndex);   // p1 뒤에 생겼으므로 p2
    }

    [Fact]
    public void 삭제는_앞_문단_뒤에_다시_넣어서_되돌린다()
    {
        var undo = BckPptViewModel.BuildUndoEdits(
            [Change(4, "구조 분석", "", ParagraphChangeKind.Clear)]);

        var edit = Assert.Single(undo);
        Assert.Equal(ParagraphEditKind.Insert, edit.Kind);
        Assert.Equal(3, edit.ParagraphIndex);   // p3 뒤 = 원래의 p4 자리
        Assert.Equal("구조 분석", edit.NewText);
    }

    [Fact]
    public void 첫문단_삭제는_맨앞_삽입으로_되돌린다()
    {
        var undo = BckPptViewModel.BuildUndoEdits(
            [Change(1, "첫 줄", "", ParagraphChangeKind.Clear)]);

        var edit = Assert.Single(undo);
        Assert.Equal(ParagraphEditKind.Insert, edit.Kind);
        Assert.Equal(0, edit.ParagraphIndex);   // 0 = 맨 앞
    }

    [Fact]
    public void 앞에서_지워지면_뒤_교체는_당겨진_인덱스를_쓴다()
    {
        // p2 삭제 + p4 교체 → 적용 후 옛 p4는 p3 자리에 있다.
        var undo = BckPptViewModel.BuildUndoEdits(
        [
            Change(2, "지운 줄", "",    ParagraphChangeKind.Clear),
            Change(4, "옛4",    "새4", ParagraphChangeKind.Replace),
        ]);

        var replace = Assert.Single(undo, e => e.Kind == ParagraphEditKind.Replace);
        Assert.Equal(3, replace.ParagraphIndex);
        Assert.Equal("옛4", replace.NewText);
    }

    [Fact]
    public void 앞에서_삽입되면_뒤_교체는_밀린_인덱스를_쓴다()
    {
        // p1 뒤 삽입 + p3 교체 → 적용 후 옛 p3은 p4 자리에 있다.
        var undo = BckPptViewModel.BuildUndoEdits(
        [
            Change(1, "",    "추가", ParagraphChangeKind.Insert),
            Change(3, "옛3", "새3",  ParagraphChangeKind.Replace),
        ]);

        var replace = Assert.Single(undo, e => e.Kind == ParagraphEditKind.Replace);
        Assert.Equal(4, replace.ParagraphIndex);
    }

    [Fact]
    public void 다른_도형끼리는_서로_인덱스에_영향을_주지_않는다()
    {
        var other = new ParagraphChange
        {
            Id = "S1-2-3", SlideIndex = 1, ShapeName = "Pentagon 69",
            ParagraphIndex = 3, OldText = "옛", NewText = "새",
            Kind = ParagraphChangeKind.Replace,
        };

        var undo = BckPptViewModel.BuildUndoEdits(
            [Change(1, "", "추가", ParagraphChangeKind.Insert), other]);

        var forOther = Assert.Single(undo, e => e.ShapeName == "Pentagon 69");
        Assert.Equal(3, forOther.ParagraphIndex);
    }
}
