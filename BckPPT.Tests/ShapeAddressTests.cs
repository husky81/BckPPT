using BckPPT.Models;

namespace BckPPT.Tests;

/// <summary>
/// 복사해 만든 카드처럼 같은 이름의 도형이 여러 개이고 그룹 안에 들어 있는 슬라이드를 가정한다.
/// 이름만으로 찾으면 첫 카드만 잡혀 나머지 카드를 고칠 수 없던 문제가 있었다.
/// </summary>
public class ShapeAddressTests
{
    private static readonly string[] Group = ["Group 171"];
    private static readonly string[] Top   = [];

    private static readonly List<(IReadOnlyList<string> GroupPath, string Name)> Slide =
    [
        (Top,   "슬라이드헤더"),
        (Group, "Text 19"),
        (Group, "Text 19"),
        (Group, "Text 19"),
        (Top,   "Text 19"),
        (Group, "Table 1"),
    ];

    [Fact]
    public void 이름만_주면_펼친_순서의_첫_도형을_고른다()
    {
        Assert.Equal(1, ShapeAddress.Parse("Text 19").Select(Slide));
        Assert.Equal(0, ShapeAddress.Parse("슬라이드헤더").Select(Slide));
    }

    [Fact]
    public void 샵_번호로_같은_이름_중_n번째를_고른다()
    {
        Assert.Equal(2, ShapeAddress.Parse("Text 19#2").Select(Slide));
        Assert.Equal(4, ShapeAddress.Parse("Text 19#4").Select(Slide));
        Assert.Equal(-1, ShapeAddress.Parse("Text 19#5").Select(Slide));
    }

    [Fact]
    public void 그룹_경로로_범위를_좁힌다()
    {
        Assert.Equal(3, ShapeAddress.Parse("Group 171/Text 19#3").Select(Slide));
        Assert.Equal(-1, ShapeAddress.Parse("Group 171/Text 19#4").Select(Slide));
        Assert.Equal(-1, ShapeAddress.Parse("Group 9/Text 19").Select(Slide));
    }

    [Fact]
    public void 표_셀_접미사를_떼어_행열로_보관한다()
    {
        var address = ShapeAddress.Parse("Group 171/Table 1 [R2C3]");
        Assert.True(address.IsTableCell);
        Assert.Equal((2, 3), (address.Row!.Value, address.Col!.Value));
        Assert.Equal(5, address.Select(Slide));
    }

    [Fact]
    public void 이름에_특수문자가_있어도_정확히_같으면_리터럴로_찾는다()
    {
        List<(IReadOnlyList<string>, string)> shapes = [(Top, "A/B"), (Top, "메모#2")];
        Assert.Equal(0, ShapeAddress.Parse("A/B").Select(shapes));
        Assert.Equal(1, ShapeAddress.Parse("메모#2").Select(shapes));
    }

    [Fact]
    public void 목록_주소는_겹칠_때만_번호를_붙인다()
    {
        Assert.Equal("Text 19", ShapeAddress.Format("Text 19", 1, 1));
        Assert.Equal("Text 19#2", ShapeAddress.Format("Text 19", 2, 3));
    }

    [Fact]
    public void 색은_COM의_BGR_순서로_바꾼다()
    {
        // 카드 번호 원의 노란색: 화면 #FFE08A → COM 0x8AE0FF
        Assert.Equal(0x8AE0FF, OleColor.FromHex("#FFE08A"));
        Assert.Equal("#FFE08A", OleColor.ToHex(0x8AE0FF));
        Assert.Throws<FormatException>(() => OleColor.FromHex("FFF"));
    }
}
