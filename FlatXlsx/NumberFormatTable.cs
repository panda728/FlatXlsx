namespace FlatXlsx;

/// <summary>
/// Interns each distinct number-format code used by the sheet and hands out the style index the
/// cell refers to it by. Filled while cells are written; styles.xml is written afterwards from
/// <see cref="Codes"/> - the same write-then-declare pattern as the shared-string table, which
/// is what lets a serializer pass a format code with no registration step at all.
/// </summary>
internal sealed class NumberFormatTable
{
    // Excel's own ceilings: roughly 200-250 custom formats per workbook depending on the
    // edition, and 255 characters per format code. Enforcing the conservative bounds here
    // keeps memory bounded when a code is (mistakenly) built from row data, and fails with a
    // named limit instead of a workbook Excel rejects.
    const int MaxFormats = 200;
    const int MaxCodeLength = 255;

    readonly Dictionary<string, int> _indexes = [];
    readonly List<string> _codes = [];

    /// <summary>Distinct codes in index order - the order styles.xml must declare them in.</summary>
    public IReadOnlyList<string> Codes => _codes;

    public int GetOrAdd(string code)
    {
        if (string.IsNullOrEmpty(code) || code.Length > MaxCodeLength || ContainsControlChar(code))
            throw new InvalidOperationException(SR.FormatCodeInvalid);

        if (_indexes.TryGetValue(code, out var index))
            return index;

        if (_codes.Count >= MaxFormats)
            throw new InvalidOperationException(SR.TooManyFormats(MaxFormats));

        index = _codes.Count;
        _indexes.Add(code, index);
        _codes.Add(code);
        return index;
    }

    // Control characters are illegal in XML 1.0 even as entities; a code carrying one could
    // only ever produce a workbook that reports itself as damaged.
    static bool ContainsControlChar(string s)
    {
        foreach (var c in s)
        {
            if (c < 0x20)
                return true;
        }
        return false;
    }
}
