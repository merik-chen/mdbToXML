using System.Data;
using System.Data.OleDb;
using MdbToXml.Models;
using MdbToXml.Utils;

namespace MdbToXml.Core;

public class SchemaExtractor
{
    private readonly AccessDatabase _database;

    public SchemaExtractor(AccessDatabase database)
    {
        _database = database;
    }

    public TableSchema ExtractTableSchema(string tableName)
    {
        var columns = ExtractColumns(tableName);
        var primaryKeys = ExtractPrimaryKeys(tableName);
        var indexes = ExtractIndexes(tableName);
        var relationships = ExtractRelationships(tableName);
        long rowCount;
        try { rowCount = _database.GetTableRowCount(tableName); }
        catch { rowCount = -1; }

        return new TableSchema(
            Name: tableName,
            Columns: columns,
            PrimaryKeyColumns: primaryKeys,
            Indexes: indexes,
            TableProperties: new List<TablePropertyInfo>(),
            Relationships: relationships,
            EstimatedRowCount: rowCount
        );
    }

    private List<ColumnInfo> ExtractColumns(string tableName)
    {
        var columnsTable = _database.GetColumnsSchema(tableName);
        var sortedRows = columnsTable.Rows.Cast<DataRow>()
            .OrderBy(r => Convert.ToInt32(r["ORDINAL_POSITION"]))
            .ToList();

        var columns = new List<ColumnInfo>();
        foreach (var row in sortedRows)
        {
            var name = row["COLUMN_NAME"].ToString()!;
            var oleDbType = (OleDbType)Convert.ToInt32(row["DATA_TYPE"]);
            var maxLength = row["CHARACTER_MAXIMUM_LENGTH"] is DBNull ? (int?)null : Convert.ToInt32(row["CHARACTER_MAXIMUM_LENGTH"]);
            var precision = row["NUMERIC_PRECISION"] is DBNull ? (int?)null : Convert.ToInt32(row["NUMERIC_PRECISION"]);
            var scale = row["NUMERIC_SCALE"] is DBNull ? (int?)null : Convert.ToInt32(row["NUMERIC_SCALE"]);
            var isNullable = row["IS_NULLABLE"] is not DBNull && Convert.ToBoolean(row["IS_NULLABLE"]);
            var defaultValue = row["COLUMN_DEFAULT"] is DBNull ? null : row["COLUMN_DEFAULT"].ToString();
            var typeInfo = TypeMapper.MapOleDbType(oleDbType, maxLength);

            columns.Add(new ColumnInfo(
                Name: name,
                OleDbType: oleDbType,
                JetType: typeInfo.JetType,
                SqlSType: typeInfo.SqlSType,
                MaxLength: maxLength,
                Precision: precision,
                Scale: scale,
                IsNullable: isNullable,
                DefaultValue: defaultValue,
                FieldProperties: new List<FieldPropertyInfo>()
            ));
        }
        return columns;
    }

    private List<string> ExtractPrimaryKeys(string tableName)
    {
        var pkTable = _database.GetPrimaryKeys(tableName);
        if (pkTable == null) return new List<string>();
        return pkTable.Rows.Cast<DataRow>()
            .Select(r => r["COLUMN_NAME"].ToString()!)
            .ToList();
    }

    private List<IndexInfo> ExtractIndexes(string tableName)
    {
        var indexTable = _database.GetIndexes(tableName);
        if (indexTable == null) return new List<IndexInfo>();

        var indexes = new List<IndexInfo>();
        var indexGroups = indexTable.Rows.Cast<DataRow>()
            .GroupBy(r => r["INDEX_NAME"].ToString()!)
            .ToList();

        foreach (var group in indexGroups)
        {
            var indexName = group.Key;
            if (indexName.StartsWith(".", StringComparison.Ordinal))
                continue;

            var firstRow = group.First();
            var keyColumns = string.Join(" ", group.Select(r => r["COLUMN_NAME"].ToString()! + " "));
            var isPrimary = firstRow["PRIMARY_KEY"] is not DBNull && Convert.ToBoolean(firstRow["PRIMARY_KEY"]);
            var isUnique = firstRow["UNIQUE"] is not DBNull && Convert.ToBoolean(firstRow["UNIQUE"]);
            var isClustered = firstRow["CLUSTERED"] is not DBNull && Convert.ToBoolean(firstRow["CLUSTERED"]);

            indexes.Add(new IndexInfo(
                Name: indexName,
                KeyColumns: keyColumns,
                IsPrimary: isPrimary,
                IsUnique: isUnique,
                IsClustered: isClustered,
                Order: "asc"
            ));
        }
        return indexes;
    }

    private List<RelationshipInfo> ExtractRelationships(string tableName)
    {
        var fkTable = _database.GetForeignKeys();
        if (fkTable == null) return new List<RelationshipInfo>();

        var relationships = new List<RelationshipInfo>();
        foreach (DataRow row in fkTable.Rows)
        {
            var pkTable = row["PK_TABLE_NAME"].ToString()!;
            var fkTableName = row["FK_TABLE_NAME"].ToString()!;
            if (pkTable == tableName || fkTableName == tableName)
            {
                relationships.Add(new RelationshipInfo(
                    Name: row["FK_NAME"]?.ToString() ?? "",
                    PrimaryTable: pkTable,
                    PrimaryColumn: row["PK_COLUMN_NAME"].ToString()!,
                    ForeignTable: fkTableName,
                    ForeignColumn: row["FK_COLUMN_NAME"].ToString()!
                ));
            }
        }
        return relationships;
    }
}
