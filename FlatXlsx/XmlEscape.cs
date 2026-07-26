using System.Buffers;
using System.Text;

namespace FlatXlsx;

/// <summary>
/// Writes strings as XML text content.
/// </summary>
/// <remarks>
/// Escaping markup characters is not sufficient on its own: XML 1.0 forbids most C0 control
/// characters outright, and no escape sequence makes them legal. Data coming from databases or
/// user input regularly carries such characters, and emitting them produces a workbook that
/// Excel rejects as corrupt. They are dropped here so that untrusted input cannot break the file.
/// </remarks>
internal static class XmlEscape
{
    static readonly byte[] _amp = Encoding.UTF8.GetBytes("&amp;");
    static readonly byte[] _lt = Encoding.UTF8.GetBytes("&lt;");
    static readonly byte[] _gt = Encoding.UTF8.GetBytes("&gt;");
    static readonly byte[] _cr = Encoding.UTF8.GetBytes("&#13;");
    static readonly byte[] _quot = Encoding.UTF8.GetBytes("&quot;");

    public static void WriteEscaped(string value, ArrayPoolBufferWriter writer)
    {
        var start = 0;
        for (var i = 0; i < value.Length; i++)
        {
            var c = value[i];
            if (c != '&' && c != '<' && c != '>' && c != '\r' && !IsForbidden(c))
                continue;

            WriteSegment(value, start, i - start, writer);
            start = i + 1;

            switch (c)
            {
                case '&': writer.Write(_amp); break;
                case '<': writer.Write(_lt); break;
                case '>': writer.Write(_gt); break;
                // A literal carriage return would be normalised to a line feed when the file is
                // read back (XML 1.0 line-end handling), silently altering the exported value.
                case '\r': writer.Write(_cr); break;
                // Forbidden characters are dropped.
            }
        }
        WriteSegment(value, start, value.Length - start, writer);
    }

    /// <summary>Writes a value for use inside a double-quoted XML attribute, escaping markup
    /// characters and the quote itself. Control characters cannot be escaped in XML 1.0 at all,
    /// so every value that reaches this method has been validated not to contain them.</summary>
    public static void WriteEscapedAttribute(string value, ArrayPoolBufferWriter writer)
    {
        var start = 0;
        for (var i = 0; i < value.Length; i++)
        {
            var c = value[i];
            if (c != '&' && c != '<' && c != '>' && c != '"')
                continue;

            WriteSegment(value, start, i - start, writer);
            start = i + 1;

            switch (c)
            {
                case '&': writer.Write(_amp); break;
                case '<': writer.Write(_lt); break;
                case '>': writer.Write(_gt); break;
                case '"': writer.Write(_quot); break;
            }
        }
        WriteSegment(value, start, value.Length - start, writer);
    }

    /// <summary>
    /// XML 1.0 permits #x9, #xA, #xD, [#x20-#xD7FF], [#xE000-#xFFFD] and code points above #xFFFF.
    /// Unpaired surrogates are not rejected here; the UTF-8 encoder replaces them with U+FFFD.
    /// </summary>
    static bool IsForbidden(char c)
        => c < 0x20
            ? c != '\t' && c != '\n' && c != '\r'
            : c == '\uFFFE' || c == '\uFFFF';

    static void WriteSegment(string value, int start, int length, ArrayPoolBufferWriter writer)
    {
        if (length <= 0)
            return;

#if NET5_0_OR_GREATER
        Encoding.UTF8.GetBytes(value.AsSpan(start, length), writer);
#else
        writer.Write(Encoding.UTF8.GetBytes(value.Substring(start, length)));
#endif
    }
}
