using System.Runtime.InteropServices;
using MdbToXml.Models;

namespace MdbToXml.Core;

/// <summary>
/// Reads Access DAO field and table properties using COM late binding.
/// Falls back gracefully if DAO is not available.
/// </summary>
public class FieldPropertyReader : IDisposable
{
    private dynamic? _dbEngine;
    private dynamic? _database;
    private bool _available;
    private bool _disposed;

    /// <summary>
    /// Known field properties that Access ExportXML includes.
    /// </summary>
    private static readonly HashSet<string> ExportableFieldProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "AllowZeroLength", "ColumnWidth", "ColumnOrder", "ColumnHidden",
        "Required", "DisplayControl", "IMEMode", "IMESentenceMode",
        "UnicodeCompression", "TextAlign", "AggregateType", "CurrencyLCID",
        "DecimalPlaces", "DefaultValue", "ResultType", "Expression"
    };

    private static readonly HashSet<string> ExportableTableProperties = new(StringComparer.OrdinalIgnoreCase)
    {
        "Orientation", "OrderByOn", "DefaultView", "Filter", "GUID",
        "DisplayViewsOnSharePointSite", "TotalsRow", "FilterOnLoad",
        "OrderByOnLoad", "HideNewField", "BackTint", "BackShade",
        "ThemeFontIndex", "AlternateBackThemeColorIndex",
        "AlternateBackTint", "AlternateBackShade",
        "DatasheetGridlinesThemeColorIndex", "DatasheetForeThemeColorIndex"
    };

    public bool IsAvailable => _available;

    public bool Initialize(string filePath)
    {
        try
        {
            var daoType = Type.GetTypeFromProgID("DAO.DBEngine.120");
            if (daoType == null)
            {
                daoType = Type.GetTypeFromProgID("DAO.DBEngine.36");
            }
            if (daoType == null)
            {
                _available = false;
                return false;
            }

            _dbEngine = Activator.CreateInstance(daoType);
            _database = _dbEngine!.OpenDatabase(filePath, false, true); // exclusive=false, readOnly=true
            _available = true;
            return true;
        }
        catch
        {
            _available = false;
            return false;
        }
    }

    public List<TablePropertyInfo> GetTableProperties(string tableName)
    {
        if (!_available || _database == null)
            return new List<TablePropertyInfo>();

        var props = new List<TablePropertyInfo>();
        try
        {
            var tableDef = _database.TableDefs[tableName];
            var properties = tableDef.Properties;
            int count = (int)properties.Count;

            for (int i = 0; i < count; i++)
            {
                try
                {
                    var prop = properties[i];
                    string name = (string)prop.Name;
                    if (ExportableTableProperties.Contains(name))
                    {
                        int type = (int)prop.Type;
                        string? value = ConvertPropertyValue(prop.Value);
                        props.Add(new TablePropertyInfo(name, type, value));
                    }
                }
                catch { /* Skip inaccessible properties */ }
            }
        }
        catch { /* Table not accessible */ }

        return props;
    }

    public List<FieldPropertyInfo> GetFieldProperties(string tableName, string fieldName)
    {
        if (!_available || _database == null)
            return new List<FieldPropertyInfo>();

        var props = new List<FieldPropertyInfo>();
        try
        {
            var field = _database.TableDefs[tableName].Fields[fieldName];
            var properties = field.Properties;
            int count = (int)properties.Count;

            for (int i = 0; i < count; i++)
            {
                try
                {
                    var prop = properties[i];
                    string name = (string)prop.Name;
                    if (ExportableFieldProperties.Contains(name))
                    {
                        int type = (int)prop.Type;
                        string? value = ConvertPropertyValue(prop.Value);
                        props.Add(new FieldPropertyInfo(name, type, value));
                    }
                }
                catch { /* Skip inaccessible properties */ }
            }
        }
        catch { /* Field not accessible */ }

        return props;
    }

    public List<IndexInfo> GetDaoIndexes(string tableName)
    {
        if (!_available || _database == null)
            return new List<IndexInfo>();

        var indexes = new List<IndexInfo>();
        try
        {
            var tableDef = _database.TableDefs[tableName];
            var idxCollection = tableDef.Indexes;
            int count = (int)idxCollection.Count;

            for (int i = 0; i < count; i++)
            {
                try
                {
                    var idx = idxCollection[i];
                    string name = (string)idx.Name;
                    bool isPrimary = (bool)idx.Primary;
                    bool isUnique = (bool)idx.Unique;
                    bool isClustered = (bool)idx.Clustered;

                    var fields = idx.Fields;
                    int fieldCount = (int)fields.Count;
                    var keyParts = new List<string>();
                    for (int j = 0; j < fieldCount; j++)
                    {
                        keyParts.Add((string)fields[j].Name + " ");
                    }

                    indexes.Add(new IndexInfo(
                        Name: name,
                        KeyColumns: string.Join("", keyParts),
                        IsPrimary: isPrimary,
                        IsUnique: isUnique,
                        IsClustered: isClustered,
                        Order: "asc"
                    ));
                }
                catch { /* Skip inaccessible indexes */ }
            }
        }
        catch { /* Table not accessible */ }

        return indexes;
    }

    /// <summary>
    /// Enrich a TableSchema with DAO properties (field properties, table properties, indexes).
    /// Returns a new TableSchema with populated properties.
    /// </summary>
    public TableSchema EnrichSchema(TableSchema schema)
    {
        if (!_available)
            return schema;

        var tableProps = GetTableProperties(schema.Name);
        var daoIndexes = GetDaoIndexes(schema.Name);

        var enrichedColumns = schema.Columns.Select(col =>
        {
            var fieldProps = GetFieldProperties(schema.Name, col.Name);
            return col with { FieldProperties = fieldProps };
        }).ToList();

        return schema with
        {
            Columns = enrichedColumns,
            TableProperties = tableProps,
            Indexes = daoIndexes.Count > 0 ? daoIndexes : schema.Indexes
        };
    }

    private static string? ConvertPropertyValue(object? value)
    {
        if (value == null) return null;
        if (value is byte[] bytes) return Convert.ToBase64String(bytes);
        return value.ToString();
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true;
            try
            {
                if (_database != null)
                {
                    _database.Close();
                    Marshal.ReleaseComObject(_database);
                }
                if (_dbEngine != null)
                    Marshal.ReleaseComObject(_dbEngine);
            }
            catch { /* Ignore COM cleanup errors */ }
            GC.SuppressFinalize(this);
        }
    }
}
