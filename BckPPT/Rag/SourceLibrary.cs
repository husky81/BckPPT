using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using BckPPT.Agent.Tools;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Data.Sqlite;
using UglyToad.PdfPig;
using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;
using A = DocumentFormat.OpenXml.Drawing;

namespace BckPPT.Rag;

/// <summary>원문 한 쪽. PDF는 실제 쪽, PPT는 슬라이드, Markdown·텍스트는 제목 기준 구간이다.</summary>
public sealed record SourcePage(int Page, string Heading, string Text);

public sealed record SourceDocument(string Name, string Path, int PageCount);

/// <summary>
/// 에이전트가 읽을 원문(보고서·RFP 등)을 쪽 단위로 쪼개 SQLite FTS5에 색인한다.
/// 한글은 조사가 붙어 낱말 단위 색인이 맞지 않으므로 trigram 토크나이저를 쓰고,
/// trigram이 못 찾는 2글자 낱말("개요" 등)은 LIKE로 보완한다. 순위는 코드에서 tf-idf로 매긴다.
/// </summary>
public sealed class SourceLibrary : ISourceLibrary
{
    /// <summary>Markdown·텍스트를 쪼갤 때 한 구간의 최대 글자 수.</summary>
    public const int MaxSectionChars = 2000;
    private const int SnippetChars = 400;
    private const int MaxReadPages = 10;

    private readonly string _connectionString;

    /// <summary>검색·읽기를 이 경로의 문서로 한정한다. null이면 색인한 문서 전체.</summary>
    private readonly HashSet<string>? _scope;

    private SourceLibrary(string connectionString, HashSet<string>? scope)
    {
        _connectionString = connectionString;
        _scope = scope;
    }

    /// <summary>색인은 공유하고 검색 대상만 첨부한 문서로 좁힌 보기. 같은 파일을 다시 첨부하면 색인을 재사용한다.</summary>
    public SourceLibrary WithScope(IEnumerable<string> paths) =>
        new(_connectionString, paths.Select(System.IO.Path.GetFullPath).ToHashSet(StringComparer.OrdinalIgnoreCase));

    private bool InScope(string path) => _scope is null || _scope.Contains(path);

    public SourceLibrary(string databasePath)
    {
        Directory.CreateDirectory(System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(databasePath))!);
        _connectionString = new SqliteConnectionStringBuilder { DataSource = databasePath, Pooling = false }.ToString();
        Initialize();
    }

    public static string DefaultDatabasePath => System.IO.Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "BckSoft", "BckPPT", "sources.db");

    public static readonly IReadOnlySet<string> SupportedExtensions =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".pdf", ".pptx", ".md", ".markdown", ".txt" };

    // ── 수집 ───────────────────────────────────────────────────

    /// <summary>파일을 색인한다. 내용이 같으면 다시 읽지 않는다. 색인한 쪽 수를 돌려준다.</summary>
    public int AddFile(string path)
    {
        path = System.IO.Path.GetFullPath(path);
        if (!SupportedExtensions.Contains(System.IO.Path.GetExtension(path)))
            throw new NotSupportedException($"지원하지 않는 원문 형식입니다: {System.IO.Path.GetFileName(path)}");
        var hash = Hash(path);
        using var connection = Open();
        using (var check = connection.CreateCommand())
        {
            check.CommandText = "SELECT page_count FROM source_docs WHERE path=$path AND hash=$hash";
            check.Parameters.AddWithValue("$path", path);
            check.Parameters.AddWithValue("$hash", hash);
            if (check.ExecuteScalar() is long count) return (int)count;
        }
        var pages = Extract(path);
        Replace(connection, path, hash, pages);
        return pages.Count;
    }

    /// <summary>폴더 아래 지원 형식 파일을 모두 색인한다. (파일 수, 쪽 수)를 돌려준다.</summary>
    public (int Files, int Pages) AddFolder(string folder)
    {
        int files = 0, pages = 0;
        foreach (var file in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories)
                     .Where(f => SupportedExtensions.Contains(System.IO.Path.GetExtension(f)))
                     .Where(f => !System.IO.Path.GetFileName(f).StartsWith("~$", StringComparison.Ordinal)))
        {
            pages += AddFile(file);
            files++;
        }
        return (files, pages);
    }

    public void Remove(string path)
    {
        using var connection = Open();
        using var transaction = connection.BeginTransaction();
        Delete(connection, transaction, System.IO.Path.GetFullPath(path));
        transaction.Commit();
    }

    public IReadOnlyList<SourceDocument> Documents()
    {
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT name, path, page_count FROM source_docs ORDER BY name";
        using var reader = command.ExecuteReader();
        var result = new List<SourceDocument>();
        while (reader.Read()) result.Add(new SourceDocument(reader.GetString(0), reader.GetString(1), reader.GetInt32(2)));
        return result.Where(d => InScope(d.Path)).ToList();
    }

    // ── 검색·읽기 (ISourceLibrary) ─────────────────────────────

    public IReadOnlyList<SourceHit> Search(string query, int limit)
    {
        var terms = Terms(query);
        if (terms.Count == 0) return [];

        using var connection = Open();
        int total = Documents().Sum(d => d.PageCount);
        if (total == 0) return [];

        // 후보: 어느 한 낱말이라도 들어 있는 쪽. 긴 낱말은 FTS로, 짧은 낱말은 LIKE로 찾는다.
        var candidates = new Dictionary<long, (string Doc, int Page, string Heading, string Text)>();
        var df = new Dictionary<string, int>();
        foreach (var term in terms)
        {
            using var command = connection.CreateCommand();
            if (term.Length >= 3)
            {
                command.CommandText = """
                    SELECT p.rowid, d.name, p.page, p.heading, p.text, d.path FROM source_fts f
                    JOIN source_pages p ON p.rowid = f.rowid JOIN source_docs d ON d.id = p.doc_id
                    WHERE source_fts MATCH $q
                    """;
                command.Parameters.AddWithValue("$q", "\"" + term.Replace("\"", "\"\"") + "\"");
            }
            else
            {
                command.CommandText = """
                    SELECT p.rowid, d.name, p.page, p.heading, p.text, d.path FROM source_pages p
                    JOIN source_docs d ON d.id = p.doc_id
                    WHERE p.text LIKE $q ESCAPE '\' OR p.heading LIKE $q ESCAPE '\'
                    """;
                command.Parameters.AddWithValue("$q", "%" + EscapeLike(term) + "%");
            }
            using var reader = command.ExecuteReader();
            int found = 0;
            while (reader.Read())
            {
                if (!InScope(reader.GetString(5))) continue;
                found++;
                candidates[reader.GetInt64(0)] = (reader.GetString(1), reader.GetInt32(2), reader.GetString(3), reader.GetString(4));
            }
            df[term] = found;
        }

        return candidates.Values
            .Select(c => (Hit: c, Score: Score(c.Heading, c.Text, terms, df, total) + PhraseBonus(c.Heading, c.Text, terms, df, total)))
            .Where(x => x.Score > 0)
            .OrderByDescending(x => x.Score).ThenBy(x => x.Hit.Doc).ThenBy(x => x.Hit.Page)
            .Take(Math.Max(1, limit))
            .Select(x => new SourceHit(x.Hit.Doc, x.Hit.Page,
                (x.Hit.Heading.Length > 0 ? $"〈{x.Hit.Heading}〉 " : "") + Snippet(x.Hit.Text, terms, df)))
            .ToList();
    }

    public string Outline(string? document)
    {
        if (string.IsNullOrWhiteSpace(document))
        {
            var docs = Documents();
            return docs.Count == 0
                ? "첨부된 원문이 없습니다."
                : "원문 목록\n" + string.Join("\n", docs.Select(d => $"- {d.Name} ({d.PageCount}쪽)"));
        }

        using var connection = Open();
        var (docId, name, pageCount) = ResolveDocument(connection, document);
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT page, heading FROM source_pages WHERE doc_id=$doc ORDER BY page";
        command.Parameters.AddWithValue("$doc", docId);
        using var reader = command.ExecuteReader();

        // 같은 제목이 이어지는 쪽은 한 줄로 묶는다(긴 절이 여러 구간으로 나뉜 경우).
        var sb = new StringBuilder($"{name} 목차 (전체 {pageCount}쪽)\n");
        int start = 0, last = 0;
        string? current = null;
        void Flush()
        {
            if (current is null) return;
            sb.Append(start == last ? $"p.{start}" : $"p.{start}–{last}").Append(' ')
              .Append(current.Length == 0 ? "(제목 없음)" : current).Append('\n');
        }
        while (reader.Read())
        {
            int page = reader.GetInt32(0);
            string heading = reader.GetString(1);
            if (heading == current) { last = page; continue; }
            Flush();
            (start, last, current) = (page, page, heading);
        }
        Flush();
        return sb.ToString();
    }

    public string Read(string document, int fromPage, int toPage)
    {
        using var connection = Open();
        var (docId, name, pageCount) = ResolveDocument(connection, document);
        fromPage = Math.Max(1, fromPage);
        toPage = Math.Min(Math.Max(fromPage, toPage), Math.Min(pageCount, fromPage + MaxReadPages - 1));
        if (fromPage > pageCount) throw new AgentToolException($"'{name}'은 {pageCount}쪽까지 있습니다.");

        using var command = connection.CreateCommand();
        command.CommandText = "SELECT page, heading, text FROM source_pages WHERE doc_id=$doc AND page BETWEEN $from AND $to ORDER BY page";
        command.Parameters.AddWithValue("$doc", docId);
        command.Parameters.AddWithValue("$from", fromPage);
        command.Parameters.AddWithValue("$to", toPage);
        using var reader = command.ExecuteReader();
        var sb = new StringBuilder($"{name} p.{fromPage}–{toPage} (전체 {pageCount}쪽)\n");
        while (reader.Read())
            sb.Append($"\n--- p.{reader.GetInt32(0)} {reader.GetString(1)}\n{reader.GetString(2)}\n");
        return sb.ToString();
    }

    // ── 쪽 나누기 ──────────────────────────────────────────────

    public static IReadOnlyList<SourcePage> Extract(string path) =>
        System.IO.Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".pdf" => ExtractPdf(path),
            ".pptx" => ExtractPptx(path),
            _ => SplitText(File.ReadAllText(path))
        };

    private static IReadOnlyList<SourcePage> ExtractPdf(string path)
    {
        using var pdf = PdfDocument.Open(path);
        var pages = new List<SourcePage>();
        foreach (var page in pdf.GetPages())
        {
            var text = Normalize(ContentOrderTextExtractor.GetText(page));
            pages.Add(new SourcePage(page.Number, FirstLine(text), text));
        }
        return pages;
    }

    private static IReadOnlyList<SourcePage> ExtractPptx(string path)
    {
        using var document = PresentationDocument.Open(path, false);
        var part = document.PresentationPart;
        var ids = part?.Presentation.SlideIdList?.Elements<DocumentFormat.OpenXml.Presentation.SlideId>() ?? [];
        var pages = new List<SourcePage>();
        foreach (var id in ids)
        {
            if (id.RelationshipId?.Value is not { } rel || part!.GetPartById(rel) is not SlidePart slide) continue;
            var paragraphs = slide.Slide.Descendants<A.Paragraph>()
                .Select(p => string.Concat(p.Descendants<A.Text>().Select(t => t.Text)).Trim())
                .Where(t => t.Length > 0)
                .ToList();
            var text = string.Join("\n", paragraphs);
            pages.Add(new SourcePage(pages.Count + 1, FirstLine(text), text));
        }
        return pages;
    }

    private static readonly Regex HeadingLine = new(@"^#{1,6}\s+(.+?)\s*#*\s*$");

    /// <summary>
    /// Markdown·텍스트를 구간으로 나눈다. 폼피드(\f)가 있으면 그것을 쪽 경계로 쓰고,
    /// 없으면 제목 줄마다 나눈 뒤 <see cref="MaxSectionChars"/>를 넘는 구간은 빈 줄 경계에서 다시 자른다.
    /// 각 구간의 heading은 가장 가까운 위 제목이다.
    /// </summary>
    public static IReadOnlyList<SourcePage> SplitText(string text)
    {
        text = text.Replace("\r\n", "\n");
        if (text.Contains('\f'))
            return text.Split('\f').Select(Normalize).Where(t => t.Length > 0)
                .Select((t, i) => new SourcePage(i + 1, FirstLine(t), t)).ToList();

        var sections = new List<(string Heading, StringBuilder Body)>();
        string heading = "";
        var body = new StringBuilder();
        foreach (var line in text.Split('\n'))
        {
            var match = HeadingLine.Match(line);
            if (match.Success)
            {
                if (body.ToString().Trim().Length > 0) sections.Add((heading, body));
                heading = match.Groups[1].Value.Trim();
                body = new StringBuilder();
            }
            body.Append(line).Append('\n');
        }
        if (body.ToString().Trim().Length > 0) sections.Add((heading, body));

        var pages = new List<SourcePage>();
        foreach (var (title, content) in sections)
            foreach (var chunk in ChunkAtBlankLines(content.ToString(), MaxSectionChars))
                pages.Add(new SourcePage(pages.Count + 1, title, chunk));
        return pages;
    }

    private static IEnumerable<string> ChunkAtBlankLines(string text, int max)
    {
        var current = new StringBuilder();
        foreach (var block in Regex.Split(text, @"\n\s*\n"))
        {
            var piece = block.Trim();
            if (piece.Length == 0) continue;
            if (current.Length > 0 && current.Length + piece.Length + 2 > max)
            {
                yield return current.ToString();
                current.Clear();
            }
            // 빈 줄 없이 긴 덩어리는 글자 수로 자른다.
            while (piece.Length > max)
            {
                if (current.Length > 0) { yield return current.ToString(); current.Clear(); }
                yield return piece[..max];
                piece = piece[max..];
            }
            if (current.Length > 0) current.Append("\n\n");
            current.Append(piece);
        }
        if (current.Length > 0) yield return current.ToString();
    }

    // ── 점수 ───────────────────────────────────────────────────

    /// <summary>낱말. 공백·구두점으로 나누고 너무 짧은(1글자) 것은 버린다.</summary>
    public static IReadOnlyList<string> Terms(string query) =>
        Regex.Split(query, @"[\s,.;:!?()\[\]{}""'·/]+")
            .Select(t => t.Trim())
            .Where(t => t.Length >= 2)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static double Score(string heading, string text, IReadOnlyList<string> terms,
        IReadOnlyDictionary<string, int> df, int total)
    {
        double score = 0;
        foreach (var term in terms)
        {
            int tf = Count(text, term);
            bool inHeading = heading.Contains(term, StringComparison.OrdinalIgnoreCase);
            if (tf == 0 && !inHeading) continue;
            double idf = Math.Log(1 + (double)total / Math.Max(1, df.GetValueOrDefault(term)));
            if (tf > 0) score += idf * (1 + Math.Log(tf));
            // 절 제목에 나오면 그 절이 주제라는 강한 신호다.
            if (inHeading) score += 2 * idf;
        }
        return score;
    }

    /// <summary>
    /// 여러 낱말 질의가 공백을 빼고 붙어서도 나오면("과업 내용" ↔ "과업내용") 낱말이 흩어진 쪽보다 앞에 둔다.
    /// </summary>
    private static double PhraseBonus(string heading, string text, IReadOnlyList<string> terms,
        IReadOnlyDictionary<string, int> df, int total)
    {
        if (terms.Count < 2) return 0;
        string phrase = string.Concat(terms);
        string Squash(string s) => Regex.Replace(s, @"\s+", "");
        bool inHeading = Squash(heading).Contains(phrase, StringComparison.OrdinalIgnoreCase);
        if (!inHeading && !Squash(text).Contains(phrase, StringComparison.OrdinalIgnoreCase)) return 0;
        double idf = terms.Sum(t => Math.Log(1 + (double)total / Math.Max(1, df.GetValueOrDefault(t))));
        return inHeading ? 2 * idf : idf;
    }

    private static string Snippet(string text, IReadOnlyList<string> terms, IReadOnlyDictionary<string, int> df)
    {
        // 가장 드문 낱말이 처음 나오는 곳을 가운데 둔다.
        int at = terms.OrderBy(t => df.GetValueOrDefault(t))
            .Select(t => text.IndexOf(t, StringComparison.OrdinalIgnoreCase))
            .FirstOrDefault(i => i >= 0, 0);
        int start = Math.Max(0, Math.Min(at - SnippetChars / 4, text.Length - SnippetChars));
        var snippet = text.Substring(start, Math.Min(SnippetChars, text.Length - start)).Replace('\n', ' ');
        return (start > 0 ? "…" : "") + snippet + (start + SnippetChars < text.Length ? "…" : "");
    }

    private static int Count(string text, string term)
    {
        int count = 0;
        for (int i = text.IndexOf(term, StringComparison.OrdinalIgnoreCase); i >= 0;
             i = text.IndexOf(term, i + term.Length, StringComparison.OrdinalIgnoreCase))
            count++;
        return count;
    }

    // ── 저장소 ─────────────────────────────────────────────────

    private (long Id, string Name, int Pages) ResolveDocument(SqliteConnection connection, string document)
    {
        var docs = new List<(long Id, string Name, string Path, int Pages)>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT id, name, path, page_count FROM source_docs";
            using var reader = command.ExecuteReader();
            while (reader.Read()) docs.Add((reader.GetInt64(0), reader.GetString(1), reader.GetString(2), reader.GetInt32(3)));
        }
        docs.RemoveAll(d => !InScope(d.Path));
        string Stem(string n) => System.IO.Path.GetFileNameWithoutExtension(n);
        var match = docs.FirstOrDefault(d => d.Name.Equals(document, StringComparison.OrdinalIgnoreCase));
        if (match.Name is null) match = docs.FirstOrDefault(d => Stem(d.Name).Equals(Stem(document), StringComparison.OrdinalIgnoreCase));
        if (match.Name is null)
        {
            var partial = docs.Where(d => d.Name.Contains(document, StringComparison.OrdinalIgnoreCase)).ToList();
            if (partial.Count == 1) match = partial[0];
        }
        if (match.Name is null)
            throw new AgentToolException($"원문 '{document}'를 찾지 못했습니다. 있는 문서: {string.Join(", ", docs.Select(d => d.Name))}");
        return (match.Id, match.Name, match.Pages);
    }

    private void Replace(SqliteConnection connection, string path, string hash, IReadOnlyList<SourcePage> pages)
    {
        using var transaction = connection.BeginTransaction();
        Delete(connection, transaction, path);
        long docId;
        using (var insert = connection.CreateCommand())
        {
            insert.Transaction = transaction;
            insert.CommandText = """
                INSERT INTO source_docs (path, name, hash, page_count, indexed_utc) VALUES ($path, $name, $hash, $count, $now);
                SELECT last_insert_rowid();
                """;
            insert.Parameters.AddWithValue("$path", path);
            insert.Parameters.AddWithValue("$name", System.IO.Path.GetFileName(path));
            insert.Parameters.AddWithValue("$hash", hash);
            insert.Parameters.AddWithValue("$count", pages.Count);
            insert.Parameters.AddWithValue("$now", DateTime.UtcNow.ToString("O"));
            docId = (long)insert.ExecuteScalar()!;
        }
        using var page = connection.CreateCommand();
        page.Transaction = transaction;
        page.CommandText = """
            INSERT INTO source_pages (doc_id, page, heading, text) VALUES ($doc, $page, $heading, $text);
            INSERT INTO source_fts (rowid, heading, text) VALUES (last_insert_rowid(), $heading, $text);
            """;
        var pDoc = page.Parameters.Add("$doc", SqliteType.Integer);
        var pPage = page.Parameters.Add("$page", SqliteType.Integer);
        var pHeading = page.Parameters.Add("$heading", SqliteType.Text);
        var pText = page.Parameters.Add("$text", SqliteType.Text);
        foreach (var p in pages)
        {
            pDoc.Value = docId;
            pPage.Value = p.Page;
            pHeading.Value = p.Heading;
            pText.Value = p.Text;
            page.ExecuteNonQuery();
        }
        transaction.Commit();
    }

    private static void Delete(SqliteConnection connection, SqliteTransaction transaction, string path)
    {
        using var command = connection.CreateCommand();
        command.Transaction = transaction;
        command.CommandText = """
            DELETE FROM source_fts WHERE rowid IN
              (SELECT p.rowid FROM source_pages p JOIN source_docs d ON d.id = p.doc_id WHERE d.path = $path);
            DELETE FROM source_pages WHERE doc_id IN (SELECT id FROM source_docs WHERE path = $path);
            DELETE FROM source_docs WHERE path = $path;
            """;
        command.Parameters.AddWithValue("$path", path);
        command.ExecuteNonQuery();
    }

    private void Initialize()
    {
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = """
            CREATE TABLE IF NOT EXISTS source_docs (
              id INTEGER PRIMARY KEY, path TEXT NOT NULL UNIQUE, name TEXT NOT NULL,
              hash TEXT NOT NULL, page_count INTEGER NOT NULL, indexed_utc TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS source_pages (
              doc_id INTEGER NOT NULL, page INTEGER NOT NULL, heading TEXT NOT NULL, text TEXT NOT NULL);
            CREATE INDEX IF NOT EXISTS source_pages_doc ON source_pages (doc_id, page);
            CREATE VIRTUAL TABLE IF NOT EXISTS source_fts USING fts5 (heading, text, tokenize = 'trigram');
            """;
        command.ExecuteNonQuery();
    }

    private SqliteConnection Open()
    {
        var connection = new SqliteConnection(_connectionString);
        connection.Open();
        return connection;
    }

    private static string Hash(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream));
    }

    private static string EscapeLike(string term) =>
        term.Replace("\\", "\\\\").Replace("%", "\\%").Replace("_", "\\_");

    private static string Normalize(string text) =>
        Regex.Replace(text.Replace("\r\n", "\n").Replace('\r', '\n'), @"[ \t]+", " ").Trim();

    // "- 5 -", "12", "5 / 40"처럼 쪽 번호만 있는 줄
    private static readonly Regex PageNumberLine = new(@"^[-–—\s\d/]+$");

    private static string FirstLine(string text)
    {
        var line = text.Split('\n').Select(l => l.Trim())
            .FirstOrDefault(l => l.Length > 0 && !PageNumberLine.IsMatch(l)) ?? "";
        return line.Length <= 80 ? line : line[..80];
    }
}
