using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using MdbToXml.Core;
using MdbToXml.Models;
using MdbToXml.Utils;

namespace MdbToXml;

public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        if (args.Length == 0 || args[0] == "-h" || args[0] == "--help")
        {
            PrintHelp();
            return 0;
        }

        string command = args[0].ToLowerInvariant();
        if (command == "info")
        {
            return RunInfo(args.Skip(1).ToArray());
        }
        else if (command == "export")
        {
            return await RunExportAsync(args.Skip(1).ToArray());
        }
        else
        {
            Console.WriteLine($"Unknown command: {command}");
            PrintHelp();
            return 1;
        }
    }

    private static void PrintHelp()
    {
        Console.WriteLine("MDB to XML/XSD Converter");
        Console.WriteLine("Usage:");
        Console.WriteLine("  mdbToXML info <database.mdb>");
        Console.WriteLine("  mdbToXML export <database.mdb> [options]");
        Console.WriteLine();
        Console.WriteLine("Export Options:");
        Console.WriteLine("  -o, --output <dir>   Output directory (default: same as db)");
        Console.WriteLine("  -t, --tables <t1,t2> Specific tables to export (comma separated)");
        Console.WriteLine("  --single-file        Merge all tables into a single XML file");
        Console.WriteLine("  --xsd-only           Generate XSD schema only, do not export data");
        Console.WriteLine("  --xml-only           Generate XML data only, do not generate XSD");
        Console.WriteLine("  --simple-xsd         Generate simplified XSD without DAO metadata");
        Console.WriteLine("  --batch-size <num>   Rows to process before flush (default: 10000)");
        Console.WriteLine("  --split-size <num>   Split output into new files every <num> rows (default: 0)");
        Console.WriteLine("  -v, --verbose        Show detailed progress output");
    }

    private static int RunInfo(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Error: Database path required.");
            return 1;
        }

        var dbPath = args[0];
        if (!File.Exists(dbPath))
        {
            Console.WriteLine($"Error: File not found - {dbPath}");
            return 1;
        }

        Console.WriteLine($"Analyzing database: {dbPath}");
        Console.WriteLine();

        try
        {
            using var db = new AccessDatabase(dbPath);
            using var propertyReader = new FieldPropertyReader();
            bool hasDao = propertyReader.Initialize(dbPath);

            Console.WriteLine($"Metadata Strategy: {(hasDao ? "DAO COM (Full metadata)" : "OleDb (Basic schema only)")}");
            Console.WriteLine();

            var tables = db.GetTableNames();
            Console.WriteLine($"Found {tables.Count} user tables:");

            var schemaExtractor = new SchemaExtractor(db);

            foreach (var tableName in tables)
            {
                try
                {
                    var schema = schemaExtractor.ExtractTableSchema(tableName);
                    if (hasDao)
                        schema = propertyReader.EnrichSchema(schema);

                    Console.WriteLine($"\n- Table: {schema.Name}");
                    Console.WriteLine($"  Estimated Rows: {(schema.EstimatedRowCount >= 0 ? schema.EstimatedRowCount.ToString("N0") : "Unknown")}");
                    Console.WriteLine($"  Columns: {schema.Columns.Count}");
                    Console.WriteLine($"  Primary Key: {(schema.PrimaryKeyColumns.Any() ? string.Join(", ", schema.PrimaryKeyColumns) : "None")}");

                    if (schema.Indexes.Count > 0)
                        Console.WriteLine($"  Indexes: {schema.Indexes.Count}");

                    if (schema.Relationships.Count > 0)
                        Console.WriteLine($"  Foreign Keys: {schema.Relationships.Count}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"\n- Table: {tableName}");
                    Console.WriteLine($"  Error reading schema: {ex.Message}");
                }
            }
            Console.WriteLine();
            return 0;
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine($"Error connecting to database: {ex.Message}");
            Console.ResetColor();
            return 1;
        }
    }

    private static async Task<int> RunExportAsync(string[] args)
    {
        if (args.Length == 0)
        {
            Console.WriteLine("Error: Database path required.");
            return 1;
        }

        var dbPath = args[0];
        if (!File.Exists(dbPath))
        {
            Console.WriteLine($"Error: File not found - {dbPath}");
            return 1;
        }

        string? outDir = null;
        string[] targetTables = Array.Empty<string>();
        bool singleFile = false;
        bool xsdOnly = false;
        bool xmlOnly = false;
        bool simpleXsd = false;
        int batchSize = 10000;
        int splitSize = 0;
        bool verbose = false;

        for (int i = 1; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-o":
                case "--output":
                    if (i + 1 < args.Length) outDir = args[++i];
                    break;
                case "-t":
                case "--tables":
                    if (i + 1 < args.Length) targetTables = args[++i].Split(',').Select(t => t.Trim()).ToArray();
                    break;
                case "--single-file": singleFile = true; break;
                case "--xsd-only": xsdOnly = true; break;
                case "--xml-only": xmlOnly = true; break;
                case "--simple-xsd": simpleXsd = true; break;
                case "--batch-size":
                    if (i + 1 < args.Length && int.TryParse(args[++i], out int bs)) batchSize = bs;
                    break;
                case "--split-size":
                    if (i + 1 < args.Length && int.TryParse(args[++i], out int ss)) splitSize = ss;
                    break;
                case "-v":
                case "--verbose": verbose = true; break;
            }
        }

        outDir ??= Path.GetDirectoryName(Path.GetFullPath(dbPath)) ?? Environment.CurrentDirectory;
        if (!Directory.Exists(outDir))
            Directory.CreateDirectory(outDir);

        Console.WriteLine($"Starting export from {Path.GetFileName(dbPath)}");
        Console.WriteLine($"Output directory: {outDir}");

        try
        {
            using var db = new AccessDatabase(dbPath);
            using var propertyReader = new FieldPropertyReader();
            bool hasDao = propertyReader.Initialize(dbPath);

            if (!hasDao && !simpleXsd)
            {
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine("Warning: DAO COM is not available. Exporting in simple XSD mode.");
                Console.ResetColor();
                simpleXsd = true;
            }

            var allTables = db.GetTableNames();
            var tablesToProcess = targetTables.Length > 0
                ? allTables.Where(t => targetTables.Contains(t, StringComparer.OrdinalIgnoreCase)).ToList()
                : allTables;

            if (tablesToProcess.Count == 0)
            {
                Console.WriteLine("No valid tables found to export.");
                return 0;
            }

            Console.WriteLine($"Processing {tablesToProcess.Count} tables...");
            Console.WriteLine();

            var schemaExtractor = new SchemaExtractor(db);
            var progress = new ProgressReporter(verbose);
            progress.SetTotalTables(tablesToProcess.Count);

            var dataExporter = new DataExporter(db, batchSize, splitSize, progress);
            var xsdGenerator = new XsdGenerator();

            int totalRowsExported = 0;

            if (singleFile)
            {
                var schemas = new List<TableSchema>();
                foreach (var tableName in tablesToProcess)
                {
                    var schema = schemaExtractor.ExtractTableSchema(tableName);
                    if (hasDao && !simpleXsd)
                        schema = propertyReader.EnrichSchema(schema);
                    schemas.Add(schema);
                }

                var baseName = Path.GetFileNameWithoutExtension(dbPath);
                var xsdFileName = $"{baseName}.xsd";
                var xmlFileName = $"{baseName}.xml";

                if (!xmlOnly)
                {
                    xsdGenerator.GenerateMultiTableXsd(schemas, Path.Combine(outDir, xsdFileName), simpleXsd);
                    if (verbose) Console.WriteLine($"  Generated {xsdFileName}");
                }

                if (!xsdOnly)
                {
                    totalRowsExported = await dataExporter.ExportAllTablesAsync(schemas, Path.Combine(outDir, xmlFileName), xsdFileName);
                }
            }
            else
            {
                var schemas = new List<TableSchema>();
                foreach (var tableName in tablesToProcess)
                {
                    var schema = schemaExtractor.ExtractTableSchema(tableName);
                    if (hasDao && !simpleXsd)
                        schema = propertyReader.EnrichSchema(schema);
                    schemas.Add(schema);
                }

                await Parallel.ForEachAsync(schemas, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount }, async (schema, ct) =>
                {
                    var cleanTableName = XmlSanitizer.SanitizeElementName(schema.Name);
                    var xsdFileName = $"{cleanTableName}.xsd";
                    var xmlFileName = $"{cleanTableName}.xml";

                    if (!xmlOnly)
                    {
                        xsdGenerator.GenerateXsd(schema, Path.Combine(outDir, xsdFileName), simpleXsd);
                    }

                    if (!xsdOnly)
                    {
                        var rows = await dataExporter.ExportTableAsync(schema.Name, Path.Combine(outDir, xmlFileName), schema, xsdFileName);
                        Interlocked.Add(ref totalRowsExported, rows);
                    }
                    else
                    {
                        progress.TableCompleted(schema.Name, 0, TimeSpan.Zero);
                    }
                });
            }

            progress.PrintSummary(totalRowsExported, tablesToProcess.Count);
            return 0;
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.Error.WriteLine($"Error during export: {ex.Message}");
            Console.ResetColor();
            return 1;
        }
    }
}
