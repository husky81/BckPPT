using System.IO;
using System.IO.Compression;
using System.Xml.Linq;

namespace BckPPT.Services;

/// <summary>
/// 디스크에 저장된 pptx에서 그림 도형의 원본 픽셀 크기를 읽는다. PowerPoint COM은 원본 픽셀을 알려 주지 않는다.
/// 쪽은 SlideID(presentation.xml의 sldId id), 도형은 Shape.Id(cNvPr id)로 맞춘다 — 쪽을 옮기거나 지워도 바뀌지 않는다.
/// 창에서 고치고 저장하지 않은 그림은 디스크에 없으니 빠진다.
/// </summary>
public static class PptxPictureReader
{
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    /// <summary>shapeIds 중 찾은 것만 (픽셀 너비, 높이)로 돌려준다. 파일을 못 읽으면 빈 사전.</summary>
    public static IReadOnlyDictionary<int, (int Width, int Height)> ReadPixelSizes(string pptxPath, int slideId, IReadOnlyCollection<int> shapeIds)
    {
        var result = new Dictionary<int, (int, int)>();
        WithMedia(pptxPath, slideId, shapeIds, (id, entry) =>
        {
            if (ReadImageSize(entry) is { } size) result[id] = size;
        });
        return result;
    }

    /// <summary>그림 도형 하나의 그림 파일 바이트(저장된 판). 없으면 null.</summary>
    public static byte[]? ReadImageBytes(string pptxPath, int slideId, int shapeId)
    {
        byte[]? bytes = null;
        WithMedia(pptxPath, slideId, [shapeId], (_, entry) =>
        {
            using var raw = entry.Open();
            using var ms = new MemoryStream();
            raw.CopyTo(ms);
            bytes = ms.ToArray();
        });
        return bytes;
    }

    /// <summary>쪽(SlideID)의 그림 도형(cNvPr id) 중 shapeIds에 든 것마다 그 그림 파일 항목으로 action을 부른다.</summary>
    private static void WithMedia(string pptxPath, int slideId, IReadOnlyCollection<int> shapeIds, Action<int, ZipArchiveEntry> action)
    {
        if (shapeIds.Count == 0 || !File.Exists(pptxPath) || !pptxPath.EndsWith(".pptx", StringComparison.OrdinalIgnoreCase)) return;
        try
        {
            // PowerPoint가 연 파일도 읽기는 된다(쓰기만 잠근다).
            using var stream = new FileStream(pptxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var zip = new ZipArchive(stream, ZipArchiveMode.Read);

            var presentation = Load(zip, "ppt/presentation.xml");
            var slideRid = presentation?.Root?.Element(P + "sldIdLst")?.Elements(P + "sldId")
                .FirstOrDefault(e => (string?)e.Attribute("id") == slideId.ToString())?.Attribute(R + "id")?.Value;
            if (slideRid is null) return;
            var slidePart = Target(zip, "ppt/presentation.xml", slideRid);
            if (slidePart is null || Load(zip, slidePart) is not { } slide) return;

            var done = new HashSet<int>();
            void Visit(XElement? cNvPr, XElement? blipFill)
            {
                if (!int.TryParse((string?)cNvPr?.Attribute("id"), out int id) || !shapeIds.Contains(id) || done.Contains(id)) return;
                var embed = blipFill?.Element(A + "blip")?.Attribute(R + "embed")?.Value;
                if (embed is null || Target(zip, slidePart, embed) is not { } media || zip.GetEntry(media) is not { } entry) return;
                done.Add(id);
                action(id, entry);
            }
            foreach (var pic in slide.Descendants(P + "pic"))
                Visit(pic.Element(P + "nvPicPr")?.Element(P + "cNvPr"), pic.Element(P + "blipFill"));
            // 그림 개체 틀(placeholder)에 넣은 그림은 p:sp의 blipFill에 있다.
            foreach (var sp in slide.Descendants(P + "sp"))
                Visit(sp.Element(P + "nvSpPr")?.Element(P + "cNvPr"), sp.Element(P + "spPr")?.Element(A + "blipFill"));
        }
        catch (Exception ex) when (ex is IOException or InvalidDataException or UnauthorizedAccessException or System.Xml.XmlException) { }
    }

    private static XDocument? Load(ZipArchive zip, string part)
    {
        var entry = zip.GetEntry(part);
        if (entry is null) return null;
        using var s = entry.Open();
        return XDocument.Load(s);
    }

    /// <summary>part의 관계 id가 가리키는 대상 경로(zip 안 경로).</summary>
    private static string? Target(ZipArchive zip, string part, string relId)
    {
        string dir = part.Contains('/') ? part[..part.LastIndexOf('/')] : "";
        string name = part[(part.LastIndexOf('/') + 1)..];
        var rels = Load(zip, (dir.Length > 0 ? dir + "/" : "") + "_rels/" + name + ".rels");
        var target = rels?.Root?.Elements(Rel + "Relationship").FirstOrDefault(e => (string?)e.Attribute("Id") == relId)?.Attribute("Target")?.Value;
        if (target is null || (string?)rels!.Root!.Elements(Rel + "Relationship").First(e => (string?)e.Attribute("Id") == relId).Attribute("TargetMode") == "External")
            return null;
        if (target.StartsWith('/')) return target.TrimStart('/');
        var parts = (dir.Length > 0 ? dir.Split('/') : []).ToList();
        foreach (var seg in target.Split('/'))
        {
            if (seg == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
            else if (seg != ".") parts.Add(seg);
        }
        return string.Join('/', parts);
    }

    private static (int, int)? ReadImageSize(ZipArchiveEntry entry)
    {
        using var raw = entry.Open();
        using var ms = new MemoryStream();
        raw.CopyTo(ms);
        ms.Position = 0;
        return ReadImageSize(ms);
    }

    /// <summary>그림 파일 머리에서 픽셀 크기를 읽는다(PNG·JPEG·GIF·BMP). 모르는 형식이면 null.</summary>
    public static (int Width, int Height)? ReadImageSize(Stream stream)
    {
        var head = new byte[32];
        int n = stream.Read(head, 0, head.Length);
        if (n >= 24 && head[0] == 0x89 && head[1] == 'P' && head[2] == 'N' && head[3] == 'G')
            return (BigEndian(head, 16), BigEndian(head, 20));
        if (n >= 10 && head[0] == 'G' && head[1] == 'I' && head[2] == 'F')
            return (head[6] | head[7] << 8, head[8] | head[9] << 8);
        if (n >= 26 && head[0] == 'B' && head[1] == 'M')
            return (BitConverter.ToInt32(head, 18), Math.Abs(BitConverter.ToInt32(head, 22)));
        if (n >= 4 && head[0] == 0xFF && head[1] == 0xD8)
        {
            stream.Position = 2;
            var buf = new byte[8];
            while (stream.Read(buf, 0, 4) == 4)
            {
                if (buf[0] != 0xFF) return null;
                int marker = buf[1], length = buf[2] << 8 | buf[3];
                // SOF0~SOF15(DHT C4·JPG C8·DAC CC 제외)에 크기가 있다.
                if (marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC)
                {
                    if (stream.Read(buf, 0, 5) != 5) return null;
                    return (buf[3] << 8 | buf[4], buf[1] << 8 | buf[2]);
                }
                if (length < 2) return null;
                stream.Position += length - 2;
            }
        }
        return null;
    }

    private static int BigEndian(byte[] b, int i) => b[i] << 24 | b[i + 1] << 16 | b[i + 2] << 8 | b[i + 3];
}
