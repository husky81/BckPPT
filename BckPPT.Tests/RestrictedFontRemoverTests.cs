using System.IO;
using BckPPT.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace BckPPT.Tests;

public sealed class RestrictedFontRemoverTests : IDisposable
{
    private readonly string _dir = Path.Combine(Path.GetTempPath(), "bckppt-font-tests-" + Guid.NewGuid().ToString("N")[..8]);

    /// <summary>EOT 머리만 흉내 낸 글꼴 데이터. 실제 글리프는 없어도 판정에는 머리만 쓴다.</summary>
    private static byte[] Eot(ushort fsType, int length = 64)
    {
        var data = new byte[length];
        BitConverter.GetBytes((uint)length).CopyTo(data, 0);
        BitConverter.GetBytes(0x00020002u).CopyTo(data, 8);
        BitConverter.GetBytes(fsType).CopyTo(data, 32);
        return data;
    }

    [Theory]
    [InlineData(0x0004, true)]    // 미리 보기·인쇄 전용 → 읽기 전용(착수보고 PPT의 G마켓 산스)
    [InlineData(0x0002, true)]    // 제한
    [InlineData(0x0008, false)]   // 편집 가능
    [InlineData(0x0000, false)]   // 설치 가능
    [InlineData(0x000C, false)]   // 편집 비트가 있으면 편집 가능
    public void 내장_권한으로_편집_차단을_판정한다(int fsType, bool blocks)
    {
        var header = Eot((ushort)fsType);
        Assert.Equal((ushort)fsType, RestrictedFontRemover.ReadFsType(header, header.Length));
        Assert.Equal(blocks, RestrictedFontRemover.BlocksEditing((ushort)fsType));
    }

    [Fact]
    public void EOT가_아니면_판정하지_않는다()
    {
        var header = Eot(0x0004);
        Assert.Null(RestrictedFontRemover.ReadFsType(header, header.Length + 1));   // 크기 불일치
        Assert.Null(RestrictedFontRemover.ReadFsType(new byte[10], 10));
    }

    [Fact]
    public void 편집을_막는_글꼴만_빼고_나머지는_남긴다()
    {
        Directory.CreateDirectory(_dir);
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "icon.pptx"));
        var path = Path.Combine(_dir, "fonts.pptx");
        File.Copy(fixture, path);
        using (var document = PresentationDocument.Open(path, true))
        {
            var part = document.PresentationPart!;
            string AddFont(ushort fsType)
            {
                var font = part.AddFontPart(FontPartType.FontData);
                using (var stream = new MemoryStream(Eot(fsType))) font.FeedData(stream);
                return part.GetIdOfPart(font);
            }
            var list = new P.EmbeddedFontList(
                new P.EmbeddedFont(new P.Font { Typeface = "G마켓 산스 Bold" },
                    new P.RegularFont { Id = AddFont(0x0004) }, new P.BoldFont { Id = AddFont(0x0004) }),
                new P.EmbeddedFont(new P.Font { Typeface = "KoPub돋움체 Bold" },
                    new P.RegularFont { Id = AddFont(0x0004) }, new P.BoldFont { Id = AddFont(0x0008) }),
                new P.EmbeddedFont(new P.Font { Typeface = "Pretendard Light" },
                    new P.RegularFont { Id = AddFont(0x0000) }));
            part.Presentation.InsertAfter(list, part.Presentation.NotesSize);
            part.Presentation.Save();
        }

        Assert.Equal(3, RestrictedFontRemover.Inspect(path).Count);
        var removed = RestrictedFontRemover.Remove(path);

        Assert.Equal(3, removed.Count);
        Assert.Empty(RestrictedFontRemover.Inspect(path));
        using var check = PresentationDocument.Open(path, false);
        var fonts = check.PresentationPart!.Presentation.EmbeddedFontList!.Elements<P.EmbeddedFont>().ToList();
        Assert.Equal(["KoPub돋움체 Bold", "Pretendard Light"], fonts.Select(f => f.Font!.Typeface!.Value));
        Assert.Single(fonts[0].Elements<P.BoldFont>());
        Assert.Empty(fonts[0].Elements<P.RegularFont>());
        Assert.Equal(2, check.PresentationPart.FontParts.Count());
    }

    [Fact]
    public void 제한_글꼴이_없으면_파일을_건드리지_않는다()
    {
        Directory.CreateDirectory(_dir);
        var fixture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "icon.pptx"));
        var path = Path.Combine(_dir, "plain.pptx");
        File.Copy(fixture, path);
        var before = File.GetLastWriteTimeUtc(path);
        Assert.Empty(RestrictedFontRemover.Remove(path));
        Assert.Equal(before, File.GetLastWriteTimeUtc(path));
    }

    public void Dispose()
    {
        try { Directory.Delete(_dir, recursive: true); } catch { }
    }
}
