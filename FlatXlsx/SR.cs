using System.Globalization;
using System.Resources;

namespace FlatXlsx;

/// <summary>
/// The library's diagnostic messages.
/// </summary>
/// <remarks>
/// Messages follow <see cref="CultureInfo.CurrentUICulture"/>, which is the convention every
/// .NET library follows and is separate from <see cref="XlsxSerializerOptions.CultureInfo"/> -
/// that one decides how <em>data</em> is formatted into the workbook, and changing the language
/// of an error message must never change the file that is produced.
///
/// English lives in the main assembly; each translation ships as a satellite assembly
/// (for example <c>ja/FlatXlsx.resources.dll</c>) that an application can simply not deploy if
/// it does not need it. A missing satellite assembly is not an error: the lookup falls back to
/// English on its own.
/// </remarks>
internal static class SR
{
    static readonly ResourceManager _resources = new("FlatXlsx.Resources.Strings", typeof(SR).Assembly);

    static string Get(string name) => _resources.GetString(name, CultureInfo.CurrentUICulture) ?? name;

    static string Format(string name, params object?[] args)
        => string.Format(CultureInfo.CurrentUICulture, Get(name), args);

    internal static string BufferAdvanceTooFar => Get(nameof(BufferAdvanceTooFar));
    internal static string InitialCapacityOutOfRange => Get(nameof(InitialCapacityOutOfRange));
    internal static string SizeHintNegative => Get(nameof(SizeHintNegative));
    internal static string MemberKindNotSupported => Get(nameof(MemberKindNotSupported));

    internal static string FormatCodeEmpty => Get(nameof(FormatCodeEmpty));

    internal static string CellTooLong(int limit, int length) => Format(nameof(CellTooLong), limit, length);
    internal static string MaxDepthReached(int depth) => Format(nameof(MaxDepthReached), depth);
    internal static string SerializerNotFound(Type type) => Format(nameof(SerializerNotFound), type);
    internal static string SheetNameInvalid(string name) => Format(nameof(SheetNameInvalid), name);
    internal static string SerializerTypeMismatch(Type declared, Type target) => Format(nameof(SerializerTypeMismatch), declared, target);
    internal static string SerializerTypeNotImplemented(Type type) => Format(nameof(SerializerTypeNotImplemented), type);
    internal static string TooManyColumns(int limit) => Format(nameof(TooManyColumns), limit);
    internal static string TooManyRows(int limit) => Format(nameof(TooManyRows), limit);
    internal static string TupleSerializerNotFound(int typeArguments) => Format(nameof(TupleSerializerNotFound), typeArguments);
}
