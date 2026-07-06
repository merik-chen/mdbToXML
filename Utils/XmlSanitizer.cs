using System.Text;

namespace MdbToXml.Utils;

public static class XmlSanitizer
{
    public static string SanitizeValue(string value)
    {
        if (string.IsNullOrEmpty(value))
            return value;

        var sb = new StringBuilder(value.Length);
        foreach (var ch in value)
        {
            if (IsLegalXmlChar(ch))
                sb.Append(ch);
        }
        return sb.ToString();
    }

    public static string SanitizeElementName(string name)
    {
        if (string.IsNullOrEmpty(name))
            return "_unnamed";

        var sb = new StringBuilder(name.Length);
        var first = name[0];
        if (char.IsLetter(first) || first == '_')
            sb.Append(first);
        else
            sb.Append('_').Append(first);

        for (int i = 1; i < name.Length; i++)
        {
            var ch = name[i];
            if (char.IsLetterOrDigit(ch) || ch == '_' || ch == '-' || ch == '.')
                sb.Append(ch);
            else
                sb.Append('_');
        }
        return sb.ToString();
    }

    private static bool IsLegalXmlChar(char ch)
    {
        return ch == 0x9 || ch == 0xA || ch == 0xD ||
               (ch >= 0x20 && ch <= 0xD7FF) ||
               (ch >= 0xE000 && ch <= 0xFFFD);
    }
}
