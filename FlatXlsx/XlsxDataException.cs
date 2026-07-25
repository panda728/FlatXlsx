namespace FlatXlsx;

/// <summary>The ways row data can fail to fit the workbook. Stable identifiers - unlike the
/// exception message, these never vary with the UI culture.</summary>
public enum XlsxDataErrorKind
{
    /// <summary>A cell value exceeds Excel's 32,767-character limit.</summary>
    CellTooLong,
    /// <summary>A row expands into more than Excel's 16,384 columns.</summary>
    TooManyColumns,
    /// <summary>The source has more rows than Excel's 1,048,576-row sheet.</summary>
    TooManyRows,
    /// <summary>The object graph nests deeper than <see cref="XlsxSerializerOptions.MaxDepth"/> -
    /// almost always a circular reference.</summary>
    MaxDepthReached,
}

/// <summary>
/// A failure caused by the row data itself: the values cannot be represented in an xlsx sheet.
/// </summary>
/// <remarks>
/// The addressee is whoever owns the data, and the properties carry what they need to act
/// without parsing the (localized) message: which rule broke (<see cref="Kind"/>), where
/// (<see cref="Row"/>/<see cref="Column"/>), and by how much (<see cref="Limit"/>/<see cref="Actual"/>).
/// Failures whose addressee is the developer - an invalid sheet name, a bad format code, a
/// missing serializer - remain plain <see cref="InvalidOperationException"/>s.
/// To collect every data error at once instead of meeting them one by one, use
/// <see cref="XlsxSerializer.Validate{T}"/>.
/// </remarks>
public sealed class XlsxDataException : InvalidOperationException
{
    public XlsxDataErrorKind Kind { get; }

    /// <summary>1-based physical row in the sheet, header row included. 0 when unknown.</summary>
    public int Row { get; }

    /// <summary>1-based column of the offending cell. 0 when the failure concerns the whole row.</summary>
    public int Column { get; }

    public long Limit { get; }
    public long Actual { get; }

    internal XlsxDataException(string message, XlsxDataErrorKind kind, int row, int column, long limit, long actual)
        : base(message)
    {
        Kind = kind;
        Row = row;
        Column = column;
        Limit = limit;
        Actual = actual;
    }
}

/// <summary>One data problem found by <see cref="XlsxSerializer.Validate{T}"/>; the same facts
/// an <see cref="XlsxDataException"/> would carry, collected instead of thrown.</summary>
public sealed record XlsxDataError(
    XlsxDataErrorKind Kind,
    int Row,
    int Column,
    long Limit,
    long Actual,
    string Message);

/// <summary>Unwinds one row during validation when it cannot be processed further (a circular
/// reference); the row loop catches it and continues with the next row.</summary>
internal sealed class RowAbortedException : Exception
{
}
