using System.Security.Cryptography;
using System.Text;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace BckPPT.Rag;

public sealed class PptIndexer(TemplateRepository repository)
{
    public async Task<(int Files, int Slides, int Skipped)> IndexFolderAsync(
        string folder, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        if (!Directory.Exists(folder)) throw new DirectoryNotFoundException(folder);
        var files = Directory.EnumerateFiles(folder, "*.pptx", SearchOption.AllDirectories)
            .Where(path => !path.Contains("~$", StringComparison.Ordinal)).ToList();
        var indexedFiles = 0; var indexedSlides = 0; var skipped = 0;
        foreach (var path in files)
        {
            cancellationToken.ThrowIfCancellationRequested();
            progress?.Report($"인덱싱: {Path.GetFileName(path)}");
            var hash = await ComputeHashAsync(path, cancellationToken);
            if (string.Equals(repository.GetFileHash(path), hash, StringComparison.Ordinal)) { skipped++; continue; }
            var slides = await Task.Run(() => ReadSlides(path, hash), cancellationToken);
            repository.ReplaceFile(path, slides);
            indexedFiles++; indexedSlides += slides.Count;
        }
        return (indexedFiles, indexedSlides, skipped);
    }

    private static IReadOnlyList<IndexedSlide> ReadSlides(string path, string hash)
    {
        using var document = PresentationDocument.Open(path, false);
        var presentationPart = document.PresentationPart ?? throw new InvalidDataException("PresentationPart가 없습니다.");
        var modified = File.GetLastWriteTimeUtc(path);
        var result = new List<IndexedSlide>();
        var slideIds = presentationPart.Presentation.SlideIdList?.Elements<P.SlideId>() ?? [];
        var index = 0;
        foreach (var slideId in slideIds)
        {
            index++;
            var relationshipId = slideId.RelationshipId?.Value;
            if (relationshipId is null || presentationPart.GetPartById(relationshipId) is not SlidePart slidePart) continue;
            var texts = slidePart.Slide.Descendants<A.Text>().Select(text => text.Text ?? "").Where(text => text.Length > 0).ToList();
            var title = texts.FirstOrDefault() ?? "";
            var fullText = string.Join("\n", texts);
            var textShapes = slidePart.Slide.Descendants<P.Shape>().Count(shape => shape.Descendants<A.Text>().Any());
            var imageCount = slidePart.Slide.Descendants<P.Picture>().Count();
            var tableCount = slidePart.Slide.Descendants<A.Table>().Count();
            var paragraphs = slidePart.Slide.Descendants<A.Paragraph>().Count();
            result.Add(new IndexedSlide(path, index, hash, modified,
                Classify(index, texts.Count, imageCount, tableCount), title, fullText,
                paragraphs, textShapes, imageCount, tableCount));
        }
        return result;
    }

    private static string Classify(int index, int textCount, int images, int tables)
    {
        if (index == 1 && textCount <= 5) return "cover";
        if (tables > 0) return "table";
        if (images >= 2 && textCount <= 8) return "visual";
        if (textCount >= 12) return "content-dense";
        return "content";
    }

    private static async Task<string> ComputeHashAsync(string path, CancellationToken token)
    {
        await using var stream = File.OpenRead(path);
        return Convert.ToHexString(await SHA256.HashDataAsync(stream, token));
    }
}
