using System.Data;
using System.Data.OleDb;

namespace MdbToXml.Core;

public class AccessDatabase : IDisposable
{
    private readonly string _filePath;
    private readonly string _connectionString;
    private bool _disposed;

    public AccessDatabase(string filePath)
    {
        _filePath = Path.GetFullPath(filePath);
        if (!File.Exists(_filePath))
            throw new FileNotFoundException($"Database file not found: {_filePath}");
        _connectionString = BuildConnectionString(_filePath);
    }

    public string FilePath => _filePath;
    public string FileName => Path.GetFileNameWithoutExtension(_filePath);

    public OleDbConnection CreateReadOnlyConnection()
    {
        var conn = new OleDbConnection(_connectionString);
        conn.Open();
        return conn;
    }

    public List<string> GetTableNames()
    {
        using var conn = CreateReadOnlyConnection();
        var schemaTable = conn.GetOleDbSchemaTable(
            OleDbSchemaGuid.Tables,
            new object[] { null, null, null, "TABLE" });

        var tables = new List<string>();
        if (schemaTable != null)
        {
            foreach (DataRow row in schemaTable.Rows)
            {
                var tableName = row["TABLE_NAME"].ToString()!;
                if (!tableName.StartsWith("MSys", StringComparison.OrdinalIgnoreCase) &&
                    !tableName.StartsWith("~", StringComparison.Ordinal))
                {
                    tables.Add(tableName);
                }
            }
        }
        return tables;
    }

    public long GetTableRowCount(string tableName)
    {
        using var conn = CreateReadOnlyConnection();
        using var cmd = new OleDbCommand($"SELECT COUNT(*) FROM [{tableName}]", conn);
        return Convert.ToInt64(cmd.ExecuteScalar());
    }

    public DataTable GetColumnsSchema(string tableName)
    {
        using var conn = CreateReadOnlyConnection();
        return conn.GetOleDbSchemaTable(
            OleDbSchemaGuid.Columns,
            new object[] { null, null, tableName, null })!;
    }

    public DataTable? GetPrimaryKeys(string tableName)
    {
        using var conn = CreateReadOnlyConnection();
        return conn.GetOleDbSchemaTable(
            OleDbSchemaGuid.Primary_Keys,
            new object[] { null, null, tableName });
    }

    public DataTable? GetIndexes(string tableName)
    {
        using var conn = CreateReadOnlyConnection();
        return conn.GetOleDbSchemaTable(
            OleDbSchemaGuid.Indexes,
            new object[] { null, null, null, null, tableName });
    }

    public DataTable? GetForeignKeys()
    {
        using var conn = CreateReadOnlyConnection();
        return conn.GetOleDbSchemaTable(OleDbSchemaGuid.Foreign_Keys, null);
    }

    private static string BuildConnectionString(string filePath)
    {
        return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={filePath};Mode=Read;";
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }
}
