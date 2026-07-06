namespace MdbToXml.Models;

public record IndexInfo(
    string Name,
    string KeyColumns,
    bool IsPrimary,
    bool IsUnique,
    bool IsClustered,
    string Order
);
