using System.Data.OleDb;

namespace MdbToXml.Models;

public record ColumnInfo(
    string Name,
    OleDbType OleDbType,
    string JetType,
    string SqlSType,
    int? MaxLength,
    int? Precision,
    int? Scale,
    bool IsNullable,
    string? DefaultValue,
    IReadOnlyList<FieldPropertyInfo> FieldProperties
);
