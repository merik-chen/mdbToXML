using System.Data.OleDb;
using System.Globalization;

namespace MdbToXml.Utils;

public static class TypeMapper
{
    public record XsdTypeInfo(
        string JetType,
        string SqlSType,
        string XsdBaseType,
        bool UseInlineRestriction
    );

    public static XsdTypeInfo MapOleDbType(OleDbType oleDbType, int? maxLength)
    {
        return oleDbType switch
        {
            OleDbType.VarWChar or OleDbType.VarChar or OleDbType.WChar or OleDbType.Char
                => new XsdTypeInfo("text", "nvarchar", "xsd:string", true),
            OleDbType.LongVarWChar or OleDbType.LongVarChar
                => new XsdTypeInfo("memo", "ntext", "xsd:string", true),
            OleDbType.Date or OleDbType.DBDate or OleDbType.DBTimeStamp
                => new XsdTypeInfo("datetime", "datetime", "xsd:dateTime", false),
            OleDbType.Single
                => new XsdTypeInfo("single", "real", "xsd:float", false),
            OleDbType.Double
                => new XsdTypeInfo("double", "float", "xsd:double", false),
            OleDbType.SmallInt
                => new XsdTypeInfo("integer", "smallint", "xsd:short", false),
            OleDbType.Integer
                => new XsdTypeInfo("longinteger", "int", "xsd:int", false),
            OleDbType.BigInt
                => new XsdTypeInfo("longinteger", "bigint", "xsd:long", false),
            OleDbType.UnsignedTinyInt or OleDbType.TinyInt
                => new XsdTypeInfo("byte", "tinyint", "xsd:unsignedByte", false),
            OleDbType.Boolean
                => new XsdTypeInfo("yesno", "bit", "xsd:boolean", false),
            OleDbType.Currency
                => new XsdTypeInfo("currency", "money", "xsd:string", true),
            OleDbType.Guid
                => new XsdTypeInfo("guid", "uniqueidentifier", "xsd:string", true),
            OleDbType.Binary or OleDbType.VarBinary
                => new XsdTypeInfo("binary", "varbinary", "xsd:base64Binary", false),
            OleDbType.LongVarBinary
                => new XsdTypeInfo("oleobject", "image", "xsd:base64Binary", false),
            OleDbType.Numeric or OleDbType.Decimal
                => new XsdTypeInfo("decimal", "decimal", "xsd:decimal", false),
            _ => new XsdTypeInfo("text", "nvarchar", "xsd:string", true)
        };
    }

    public static int GetMaxLengthForXsd(OleDbType oleDbType, int? maxLength)
    {
        if (oleDbType is OleDbType.LongVarWChar or OleDbType.LongVarChar)
            return 536870910;
        return maxLength ?? 255;
    }

    public static string FormatValue(object value, OleDbType oleDbType)
    {
        if (value is null or DBNull)
            return string.Empty;

        return oleDbType switch
        {
            OleDbType.Date or OleDbType.DBDate or OleDbType.DBTimeStamp
                => ((DateTime)value).ToString("yyyy-MM-ddTHH:mm:ss"),
            OleDbType.Boolean
                => ((bool)value) ? "1" : "0",
            OleDbType.Single
                => ((float)value).ToString(CultureInfo.InvariantCulture),
            OleDbType.Double
                => ((double)value).ToString(CultureInfo.InvariantCulture),
            OleDbType.Decimal or OleDbType.Numeric or OleDbType.Currency
                => ((decimal)value).ToString(CultureInfo.InvariantCulture),
            OleDbType.LongVarBinary or OleDbType.Binary or OleDbType.VarBinary
                => Convert.ToBase64String((byte[])value),
            _ => value.ToString() ?? string.Empty
        };
    }
}
