using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace BckPPT.PowerPoint;

/// <summary>
/// "미리 보기·인쇄 전용"이나 "제한" 권한으로 내장된 글꼴을 PPTX에서 뺀다.
/// 이런 글꼴이 하나라도 있으면 PowerPoint는 문서를 읽기 전용으로 열어 COM 편집이 모두
/// "Presentation cannot be modified"로 실패한다(2026-10-06 착수보고 PPT의 G마켓 산스·KoPub돋움체).
/// PowerPoint의 "제한된 글꼴 제거"와 같은 처리이며, 빠진 글꼴은 설치된 글꼴로 대체되어 보인다.
/// 작업 사본에만 쓴다.
/// </summary>
public static class RestrictedFontRemover
{
    // OS/2 fsType 내장 권한 비트
    private const ushort RestrictedLicense = 0x0002;
    private const ushort PreviewAndPrint   = 0x0004;
    private const ushort Editable          = 0x0008;

    public sealed record RemovedFont(string Typeface, string Style, ushort FsType);

    /// <summary>
    /// 내장 글꼴(.fntdata, EOT 형식) 머리의 fsType을 읽는다. EOT가 아니면 null.
    /// EOT 머리: EOTSize(4) FontDataSize(4) Version(4) Flags(4) PANOSE(10) Charset(1) Italic(1) Weight(4) fsType(2)
    /// </summary>
    public static ushort? ReadFsType(ReadOnlySpan<byte> header, long length)
    {
        if (header.Length < 34) return null;
        uint eotSize = BitConverter.ToUInt32(header[..4]);
        uint version = BitConverter.ToUInt32(header[8..12]);
        if (eotSize != length || version is not (0x00010000 or 0x00020001 or 0x00020002)) return null;
        return BitConverter.ToUInt16(header[32..34]);
    }

    /// <summary>편집 권한 없이 제한·미리 보기 권한만 있으면 문서를 읽기 전용으로 만든다.</summary>
    public static bool BlocksEditing(ushort fsType) =>
        (fsType & Editable) == 0 && (fsType & (RestrictedLicense | PreviewAndPrint)) != 0;

    /// <summary>편집을 막는 내장 글꼴을 지우고 지운 목록을 돌려준다. 없으면 파일을 건드리지 않는다.</summary>
    public static IReadOnlyList<RemovedFont> Remove(string pptxPath)
    {
        if (Inspect(pptxPath).Count == 0) return [];

        var removed = new List<RemovedFont>();
        using var document = PresentationDocument.Open(pptxPath, true);
        var part = document.PresentationPart!;
        var list = part.Presentation.EmbeddedFontList;
        if (list is null) return removed;

        foreach (var font in list.Elements<P.EmbeddedFont>().ToList())
        {
            string typeface = font.Font?.Typeface?.Value ?? "(이름 없음)";
            foreach (var style in font.ChildElements.OfType<P.EmbeddedFontDataIdType>().ToList())
            {
                if (style.Id?.Value is not { } rid || part.GetPartById(rid) is not { } fontPart) continue;
                if (ReadFsType(fontPart) is not { } fsType || !BlocksEditing(fsType)) continue;
                style.Remove();
                // 같은 글꼴 파일을 다른 스타일이 함께 가리킬 수 있으므로 관계만 끊고, 남은 참조가 없으면 파트가 지워진다.
                part.DeletePart(rid);
                removed.Add(new RemovedFont(typeface, style.LocalName, fsType));
            }
            if (!font.ChildElements.OfType<P.EmbeddedFontDataIdType>().Any()) font.Remove();
        }
        if (!list.HasChildren) list.Remove();
        part.Presentation.Save();
        return removed;
    }

    /// <summary>편집을 막는 내장 글꼴 목록(읽기만).</summary>
    public static IReadOnlyList<RemovedFont> Inspect(string pptxPath)
    {
        var found = new List<RemovedFont>();
        using var document = PresentationDocument.Open(pptxPath, false);
        var part = document.PresentationPart;
        if (part?.Presentation.EmbeddedFontList is not { } list) return found;
        foreach (var font in list.Elements<P.EmbeddedFont>())
            foreach (var style in font.ChildElements.OfType<P.EmbeddedFontDataIdType>())
            {
                if (style.Id?.Value is not { } rid || part.GetPartById(rid) is not { } fontPart) continue;
                if (ReadFsType(fontPart) is { } fsType && BlocksEditing(fsType))
                    found.Add(new RemovedFont(font.Font?.Typeface?.Value ?? "(이름 없음)", style.LocalName, fsType));
            }
        return found;
    }

    private static ushort? ReadFsType(OpenXmlPart fontPart)
    {
        using var stream = fontPart.GetStream(FileMode.Open, FileAccess.Read);
        var header = new byte[34];
        int read = 0;
        while (read < header.Length && stream.Read(header, read, header.Length - read) is > 0 and var n) read += n;
        return read < header.Length ? null : ReadFsType(header, stream.Length);
    }
}
