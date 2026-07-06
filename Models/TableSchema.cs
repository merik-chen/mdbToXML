namespace MdbToXml.Models;

public record TableSchema(
    string Name,
    IReadOnlyList<ColumnInfo> Columns,
    IReadOnlyList<string> PrimaryKeyColumns,
    IReadOnlyList<IndexInfo> Indexes,
    IReadOnlyList<TablePropertyInfo> TableProperties,
    IReadOnlyList<RelationshipInfo> Relationships,
    long EstimatedRowCount
);
