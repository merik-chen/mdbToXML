using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Text;
using System.Xml;
using MdbToXml.Models;
using MdbToXml.Utils;

namespace MdbToXml.Core;

public class DataExporter
{
    private readonly AccessDatabase _database;
    private readonly int _batchSize;
    private readonly int _splitSize;
    private readonly ProgressReporter _progress;

    public DataExporter(AccessDatabase database, int batchSize, int splitSize, ProgressReporter progress)
    {
        _database = database;
        _batchSize = batchSize;
        _splitSize = splitSize;
        _progress = progress;
    }

    public async Task<int> ExportTableAsync(string tableName, string outputPath, TableSchema schema, string? xsdFileName = null)
    {
        var settings = new XmlWriterSettings
        {
            Async = true,
            Indent = true,
            IndentChars = "  ",
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false
        };

        var sw = Stopwatch.StartNew();

        int rowCount = 0;
        int fileIndex = 1;
        int rowsInCurrentFile = 0;
        bool hasSplit = false;
        XmlWriter? writer = null;
        FileStream? stream = null;

        async Task OpenWriterAsync()
        {
            string path = outputPath;
            if (hasSplit)
            {
                string dir = Path.GetDirectoryName(outputPath) ?? "";
                string name = Path.GetFileNameWithoutExtension(outputPath);
                string ext = Path.GetExtension(outputPath);
                path = Path.Combine(dir, $"{name}_{fileIndex:D3}{ext}");
            }

            stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 65536);
            writer = XmlWriter.Create(stream, settings);

            await writer.WriteStartDocumentAsync();
            await writer.WriteStartElementAsync(null, "dataroot", null);
            writer.WriteAttributeString("xmlns", "od", null, "urn:schemas-microsoft-com:officedata");
            writer.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
            writer.WriteAttributeString("xsi", "noNamespaceSchemaLocation", null, xsdFileName ?? $"{tableName}.xsd");
            writer.WriteAttributeString("generated", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"));
        }

        async Task CloseWriterAsync()
        {
            if (writer != null)
            {
                await writer.WriteFullEndElementAsync(); // </dataroot>
                await writer.WriteEndDocumentAsync();
                await writer.FlushAsync();
                await writer.DisposeAsync();
                if (stream != null)
                    await stream.DisposeAsync();
                writer = null;
                stream = null;
            }
        }

        await OpenWriterAsync();

        // Build explicit SELECT query to ensure column order matches schema
        var columnNames = string.Join(", ", schema.Columns.Select(c => $"[{c.Name}]"));
        using var conn = _database.CreateReadOnlyConnection();
        using var cmd = new OleDbCommand($"SELECT {columnNames} FROM [{tableName}]", conn);
        cmd.CommandTimeout = 0; // No timeout for large tables
        using var reader = cmd.ExecuteReader(CommandBehavior.SequentialAccess);

        bool hasMoreRows = true;
        while (true)
        {
            try
            {
                hasMoreRows = reader.Read();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"\n  [Warning] Database read error at row {rowCount} in table {tableName}. The database may contain corrupted rows or unsupported objects. Stopping export for this table. Error: {ex.Message}");
                Console.ResetColor();
                break; // Break the loop but allow XML to close cleanly
            }

            if (!hasMoreRows)
                break;

            if (_splitSize > 0 && rowsInCurrentFile >= _splitSize)
            {
                await CloseWriterAsync();

                if (!hasSplit)
                {
                    string dir = Path.GetDirectoryName(outputPath) ?? "";
                    string name = Path.GetFileNameWithoutExtension(outputPath);
                    string ext = Path.GetExtension(outputPath);
                    string newFirstPath = Path.Combine(dir, $"{name}_001{ext}");
                    if (File.Exists(outputPath))
                    {
                        if (File.Exists(newFirstPath)) File.Delete(newFirstPath);
                        File.Move(outputPath, newFirstPath);
                    }
                    hasSplit = true;
                }

                fileIndex++;
                rowsInCurrentFile = 0;
                await OpenWriterAsync();
            }

            await writer!.WriteStartElementAsync(null, tableName, null);

            for (int i = 0; i < schema.Columns.Count; i++)
            {
                var col = schema.Columns[i];
                await writer.WriteStartElementAsync(null, col.Name, null);

                if (!reader.IsDBNull(i))
                {
                    object? rawValue = null;
                    try
                    {
                        rawValue = reader.GetValue(i);
                    }
                    catch (Exception)
                    {
                        // Some OLE Objects or corrupted fields throw "The provider could not determine the Object value."
                        // We swallow the exception and treat the value as empty.
                    }

                    if (rawValue != null)
                    {
                        var formatted = TypeMapper.FormatValue(rawValue, col.OleDbType);
                        var sanitized = XmlSanitizer.SanitizeValue(formatted);
                        if (!string.IsNullOrEmpty(sanitized))
                            await writer.WriteStringAsync(sanitized);
                    }
                }

                // WriteFullEndElement ensures <tag></tag> instead of <tag/>
                await writer.WriteFullEndElementAsync();
            }

            await writer.WriteFullEndElementAsync(); // </TableName>

            rowCount++;
            rowsInCurrentFile++;

            if (rowCount % _batchSize == 0)
            {
                await writer.FlushAsync();
                _progress.Report(tableName, rowCount);
            }
        }

        await CloseWriterAsync();

        sw.Stop();
        _progress.TableCompleted(tableName, rowCount, sw.Elapsed);
        return rowCount;
    }

    /// <summary>
    /// Export multiple tables into a single XML file (or split files spanning multiple tables).
    /// </summary>
    public async Task<int> ExportAllTablesAsync(IReadOnlyList<TableSchema> schemas, string outputPath, string? xsdFileName = null)
    {
        var settings = new XmlWriterSettings
        {
            Async = true,
            Indent = true,
            IndentChars = "  ",
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false
        };

        var totalRows = 0;
        int fileIndex = 1;
        int rowsInCurrentFile = 0;
        bool hasSplit = false;
        XmlWriter? writer = null;
        FileStream? stream = null;

        async Task OpenWriterAsync()
        {
            string path = outputPath;
            if (hasSplit)
            {
                string dir = Path.GetDirectoryName(outputPath) ?? "";
                string name = Path.GetFileNameWithoutExtension(outputPath);
                string ext = Path.GetExtension(outputPath);
                path = Path.Combine(dir, $"{name}_{fileIndex:D3}{ext}");
            }

            stream = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, bufferSize: 65536);
            writer = XmlWriter.Create(stream, settings);

            await writer.WriteStartDocumentAsync();
            await writer.WriteStartElementAsync(null, "dataroot", null);
            writer.WriteAttributeString("xmlns", "od", null, "urn:schemas-microsoft-com:officedata");
            writer.WriteAttributeString("xmlns", "xsi", null, "http://www.w3.org/2001/XMLSchema-instance");
            if (xsdFileName != null)
                writer.WriteAttributeString("xsi", "noNamespaceSchemaLocation", null, xsdFileName);
            writer.WriteAttributeString("generated", DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss"));
        }

        async Task CloseWriterAsync()
        {
            if (writer != null)
            {
                await writer.WriteFullEndElementAsync(); // </dataroot>
                await writer.WriteEndDocumentAsync();
                await writer.FlushAsync();
                await writer.DisposeAsync();
                if (stream != null)
                    await stream.DisposeAsync();
                writer = null;
                stream = null;
            }
        }

        await OpenWriterAsync();

        foreach (var schema in schemas)
        {
            var sw = Stopwatch.StartNew();
            int rowCount = 0;

            using var conn = _database.CreateReadOnlyConnection();
            var columnNames = string.Join(", ", schema.Columns.Select(c => $"[{c.Name}]"));
            using var cmd = new OleDbCommand($"SELECT {columnNames} FROM [{schema.Name}]", conn);
            cmd.CommandTimeout = 0;
            using var reader = cmd.ExecuteReader(CommandBehavior.SequentialAccess);

            bool hasMoreRows = true;
            while (true)
            {
                try
                {
                    hasMoreRows = reader.Read();
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine($"\n  [Warning] Database read error at row {rowCount} in table {schema.Name}. The database may contain corrupted rows or unsupported objects. Stopping export for this table. Error: {ex.Message}");
                    Console.ResetColor();
                    break;
                }

                if (!hasMoreRows)
                    break;

                if (_splitSize > 0 && rowsInCurrentFile >= _splitSize)
                {
                    await CloseWriterAsync();

                    if (!hasSplit)
                    {
                        string dir = Path.GetDirectoryName(outputPath) ?? "";
                        string name = Path.GetFileNameWithoutExtension(outputPath);
                        string ext = Path.GetExtension(outputPath);
                        string newFirstPath = Path.Combine(dir, $"{name}_001{ext}");
                        if (File.Exists(outputPath))
                        {
                            if (File.Exists(newFirstPath)) File.Delete(newFirstPath);
                            File.Move(outputPath, newFirstPath);
                        }
                        hasSplit = true;
                    }

                    fileIndex++;
                    rowsInCurrentFile = 0;
                    await OpenWriterAsync();
                }

                await writer!.WriteStartElementAsync(null, schema.Name, null);
                for (int i = 0; i < schema.Columns.Count; i++)
                {
                    var col = schema.Columns[i];
                    await writer.WriteStartElementAsync(null, col.Name, null);
                    if (!reader.IsDBNull(i))
                    {
                        object? rawValue = null;
                        try
                        {
                            rawValue = reader.GetValue(i);
                        }
                        catch (Exception)
                        {
                            // Swallow read errors (e.g. uninitialized OLE Objects)
                        }

                        if (rawValue != null)
                        {
                            var formatted = TypeMapper.FormatValue(rawValue, col.OleDbType);
                            var sanitized = XmlSanitizer.SanitizeValue(formatted);
                            if (!string.IsNullOrEmpty(sanitized))
                                await writer.WriteStringAsync(sanitized);
                        }
                    }
                    await writer.WriteFullEndElementAsync();
                }
                await writer.WriteFullEndElementAsync();

                rowCount++;
                rowsInCurrentFile++;

                if (rowCount % _batchSize == 0)
                {
                    await writer.FlushAsync();
                    _progress.Report(schema.Name, rowCount);
                }
            }

            sw.Stop();
            _progress.TableCompleted(schema.Name, rowCount, sw.Elapsed);
            totalRows += rowCount;
        }

        await CloseWriterAsync();
        return totalRows;
    }
}
