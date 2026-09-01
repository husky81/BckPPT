using Microsoft.Data.Sqlite;
using System.IO;

namespace BckPPT.Rag;

public sealed class TemplateRepository
{
    private readonly string _connectionString;

    public TemplateRepository(string databasePath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(databasePath)!);
        _connectionString = new SqliteConnectionStringBuilder { DataSource = databasePath }.ToString();
        Initialize();
    }

    public string? GetFileHash(string path)
    {
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT file_hash FROM indexed_slides WHERE file_path=$path LIMIT 1";
        command.Parameters.AddWithValue("$path", path);
        return command.ExecuteScalar() as string;
    }

    public void ReplaceFile(string path, IReadOnlyList<IndexedSlide> slides)
    {
        using var connection = Open();
        using var transaction = connection.BeginTransaction();
        using (var delete = connection.CreateCommand())
        {
            delete.Transaction = transaction;
            delete.CommandText = "DELETE FROM indexed_slides WHERE file_path=$path";
            delete.Parameters.AddWithValue("$path", path);
            delete.ExecuteNonQuery();
        }
        foreach (var slide in slides)
        {
            using var insert = connection.CreateCommand();
            insert.Transaction = transaction;
            insert.CommandText = """
                INSERT INTO indexed_slides
                (file_path, slide_index, file_hash, modified_utc, category, title, text,
                 paragraph_count, text_shape_count, image_count, table_count)
                VALUES ($path,$index,$hash,$modified,$category,$title,$text,$paragraphs,$textShapes,$images,$tables)
                """;
            insert.Parameters.AddWithValue("$path", slide.FilePath);
            insert.Parameters.AddWithValue("$index", slide.SlideIndex);
            insert.Parameters.AddWithValue("$hash", slide.FileHash);
            insert.Parameters.AddWithValue("$modified", slide.FileModifiedUtc.ToString("O"));
            insert.Parameters.AddWithValue("$category", slide.Category);
            insert.Parameters.AddWithValue("$title", slide.Title);
            insert.Parameters.AddWithValue("$text", slide.Text);
            insert.Parameters.AddWithValue("$paragraphs", slide.ParagraphCount);
            insert.Parameters.AddWithValue("$textShapes", slide.TextShapeCount);
            insert.Parameters.AddWithValue("$images", slide.ImageCount);
            insert.Parameters.AddWithValue("$tables", slide.TableCount);
            insert.ExecuteNonQuery();
        }
        transaction.Commit();
    }

    public IReadOnlyList<IndexedSlide> ReadAll()
    {
        var result = new List<IndexedSlide>();
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT file_path,slide_index,file_hash,modified_utc,category,title,text,paragraph_count,text_shape_count,image_count,table_count FROM indexed_slides";
        using var reader = command.ExecuteReader();
        while (reader.Read()) result.Add(new IndexedSlide(
            reader.GetString(0), reader.GetInt32(1), reader.GetString(2), DateTime.Parse(reader.GetString(3)),
            reader.GetString(4), reader.GetString(5), reader.GetString(6), reader.GetInt32(7),
            reader.GetInt32(8), reader.GetInt32(9), reader.GetInt32(10)));
        return result;
    }

    public void DeleteAll()
    {
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = "DELETE FROM indexed_slides";
        command.ExecuteNonQuery();
    }

    private SqliteConnection Open() { var connection = new SqliteConnection(_connectionString); connection.Open(); return connection; }
    private void Initialize()
    {
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = """
            CREATE TABLE IF NOT EXISTS indexed_slides (
              file_path TEXT NOT NULL, slide_index INTEGER NOT NULL, file_hash TEXT NOT NULL,
              modified_utc TEXT NOT NULL, category TEXT NOT NULL, title TEXT NOT NULL, text TEXT NOT NULL,
              paragraph_count INTEGER NOT NULL, text_shape_count INTEGER NOT NULL,
              image_count INTEGER NOT NULL, table_count INTEGER NOT NULL,
              PRIMARY KEY(file_path, slide_index));
            CREATE INDEX IF NOT EXISTS idx_indexed_slides_category ON indexed_slides(category);
            """;
        command.ExecuteNonQuery();
    }
}
