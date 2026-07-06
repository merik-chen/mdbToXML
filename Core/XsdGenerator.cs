using System.Text;
using System.Xml;
using MdbToXml.Models;
using MdbToXml.Utils;

namespace MdbToXml.Core;

public class XsdGenerator
{
    private const string XsdNs = "http://www.w3.org/2001/XMLSchema";
    private const string OdNs = "urn:schemas-microsoft-com:officedata";

    public void GenerateXsd(TableSchema schema, string outputPath, bool simpleMode = false)
    {
        var settings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false
        };

        using var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = XmlWriter.Create(stream, settings);

        // <xsd:schema>
        writer.WriteStartElement("xsd", "schema", XsdNs);
        writer.WriteAttributeString("xmlns", "od", null, OdNs);

        // Root element: dataroot
        WriteDatarootElement(writer, schema.Name);

        // Table element with full annotations
        WriteTableElement(writer, schema, simpleMode);

        writer.WriteEndElement(); // </xsd:schema>
        writer.Flush();
    }

    /// <summary>
    /// Generate a combined XSD for multiple tables.
    /// </summary>
    public void GenerateMultiTableXsd(IReadOnlyList<TableSchema> schemas, string outputPath, bool simpleMode = false)
    {
        var settings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            Encoding = new UTF8Encoding(false),
            OmitXmlDeclaration = false
        };

        using var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var writer = XmlWriter.Create(stream, settings);

        writer.WriteStartElement("xsd", "schema", XsdNs);
        writer.WriteAttributeString("xmlns", "od", null, OdNs);

        // dataroot references all tables
        writer.WriteStartElement("xsd", "element", XsdNs);
        writer.WriteAttributeString("name", "dataroot");
        writer.WriteStartElement("xsd", "complexType", XsdNs);
        writer.WriteStartElement("xsd", "sequence", XsdNs);
        foreach (var schema in schemas)
        {
            writer.WriteStartElement("xsd", "element", XsdNs);
            writer.WriteAttributeString("ref", schema.Name);
            writer.WriteAttributeString("minOccurs", "0");
            writer.WriteAttributeString("maxOccurs", "unbounded");
            writer.WriteEndElement();
        }
        writer.WriteEndElement(); // </xsd:sequence>
        writer.WriteStartElement("xsd", "attribute", XsdNs);
        writer.WriteAttributeString("name", "generated");
        writer.WriteAttributeString("type", "xsd:dateTime");
        writer.WriteEndElement();
        writer.WriteEndElement(); // </xsd:complexType>
        writer.WriteEndElement(); // </xsd:element>

        foreach (var schema in schemas)
            WriteTableElement(writer, schema, simpleMode);

        writer.WriteEndElement(); // </xsd:schema>
        writer.Flush();
    }

    private void WriteDatarootElement(XmlWriter writer, string tableName)
    {
        writer.WriteStartElement("xsd", "element", XsdNs);
        writer.WriteAttributeString("name", "dataroot");

        writer.WriteStartElement("xsd", "complexType", XsdNs);
        writer.WriteStartElement("xsd", "sequence", XsdNs);

        writer.WriteStartElement("xsd", "element", XsdNs);
        writer.WriteAttributeString("ref", tableName);
        writer.WriteAttributeString("minOccurs", "0");
        writer.WriteAttributeString("maxOccurs", "unbounded");
        writer.WriteEndElement(); // </xsd:element ref>

        writer.WriteEndElement(); // </xsd:sequence>

        writer.WriteStartElement("xsd", "attribute", XsdNs);
        writer.WriteAttributeString("name", "generated");
        writer.WriteAttributeString("type", "xsd:dateTime");
        writer.WriteEndElement(); // </xsd:attribute>

        writer.WriteEndElement(); // </xsd:complexType>
        writer.WriteEndElement(); // </xsd:element dataroot>
    }

    private void WriteTableElement(XmlWriter writer, TableSchema schema, bool simpleMode)
    {
        writer.WriteStartElement("xsd", "element", XsdNs);
        writer.WriteAttributeString("name", schema.Name);

        // Table-level annotations (indexes + table properties)
        if (!simpleMode && (schema.Indexes.Count > 0 || schema.TableProperties.Count > 0))
        {
            writer.WriteStartElement("xsd", "annotation", XsdNs);
            writer.WriteStartElement("xsd", "appinfo", XsdNs);

            foreach (var idx in schema.Indexes)
            {
                writer.WriteStartElement("od", "index", OdNs);
                writer.WriteAttributeString("index-name", idx.Name);
                writer.WriteAttributeString("index-key", idx.KeyColumns);
                writer.WriteAttributeString("primary", idx.IsPrimary ? "yes" : "no");
                writer.WriteAttributeString("unique", idx.IsUnique ? "yes" : "no");
                writer.WriteAttributeString("clustered", idx.IsClustered ? "yes" : "no");
                writer.WriteAttributeString("order", idx.Order);
                writer.WriteEndElement();
            }

            foreach (var prop in schema.TableProperties)
            {
                writer.WriteStartElement("od", "tableProperty", OdNs);
                writer.WriteAttributeString("name", prop.Name);
                writer.WriteAttributeString("type", prop.Type.ToString());
                writer.WriteAttributeString("value", prop.Value ?? "");
                writer.WriteEndElement();
            }

            writer.WriteEndElement(); // </xsd:appinfo>
            writer.WriteEndElement(); // </xsd:annotation>
        }

        // <xsd:complexType> <xsd:sequence>
        writer.WriteStartElement("xsd", "complexType", XsdNs);
        writer.WriteStartElement("xsd", "sequence", XsdNs);

        foreach (var col in schema.Columns)
            WriteColumnElement(writer, col, simpleMode);

        writer.WriteEndElement(); // </xsd:sequence>
        writer.WriteEndElement(); // </xsd:complexType>
        writer.WriteEndElement(); // </xsd:element tableName>
    }

    private void WriteColumnElement(XmlWriter writer, ColumnInfo col, bool simpleMode)
    {
        var typeInfo = TypeMapper.MapOleDbType(col.OleDbType, col.MaxLength);

        writer.WriteStartElement("xsd", "element", XsdNs);
        writer.WriteAttributeString("name", col.Name);
        writer.WriteAttributeString("minOccurs", "0");

        // od: type attributes
        writer.WriteAttributeString("od", "jetType", OdNs, typeInfo.JetType);
        writer.WriteAttributeString("od", "sqlSType", OdNs, typeInfo.SqlSType);

        // For non-restriction types, set type attribute directly
        if (!typeInfo.UseInlineRestriction)
            writer.WriteAttributeString("type", typeInfo.XsdBaseType);

        // Field-level annotations (od:fieldProperty)
        if (!simpleMode && col.FieldProperties.Count > 0)
        {
            writer.WriteStartElement("xsd", "annotation", XsdNs);
            writer.WriteStartElement("xsd", "appinfo", XsdNs);

            foreach (var prop in col.FieldProperties)
            {
                writer.WriteStartElement("od", "fieldProperty", OdNs);
                writer.WriteAttributeString("name", prop.Name);
                writer.WriteAttributeString("type", prop.Type.ToString());
                writer.WriteAttributeString("value", prop.Value ?? "");
                writer.WriteEndElement();
            }

            writer.WriteEndElement(); // </xsd:appinfo>
            writer.WriteEndElement(); // </xsd:annotation>
        }

        // Inline restriction for string types
        if (typeInfo.UseInlineRestriction)
        {
            var maxLen = TypeMapper.GetMaxLengthForXsd(col.OleDbType, col.MaxLength);
            writer.WriteStartElement("xsd", "simpleType", XsdNs);
            writer.WriteStartElement("xsd", "restriction", XsdNs);
            writer.WriteAttributeString("base", typeInfo.XsdBaseType);
            writer.WriteStartElement("xsd", "maxLength", XsdNs);
            writer.WriteAttributeString("value", maxLen.ToString());
            writer.WriteEndElement(); // </xsd:maxLength>
            writer.WriteEndElement(); // </xsd:restriction>
            writer.WriteEndElement(); // </xsd:simpleType>
        }

        writer.WriteEndElement(); // </xsd:element>
    }
}
