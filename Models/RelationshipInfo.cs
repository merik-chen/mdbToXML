namespace MdbToXml.Models;

public record RelationshipInfo(
    string Name,
    string PrimaryTable,
    string PrimaryColumn,
    string ForeignTable,
    string ForeignColumn
);
