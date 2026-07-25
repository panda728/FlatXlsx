namespace FlatXlsx;

/// <summary>
/// Interns each distinct number-format code used by the sheet and hands out the style index the
/// cell refers to it by. Filled while cells are written; styles.xml is written afterwards from
/// <see cref="Codes"/> - the same write-then-declare pattern as the shared-string table, which
/// is what lets a serializer pass a format code with no registration step at all.
/// </summary>
internal sealed class NumberFormatTable
{
    readonly Dictionary<string, int> _indexes = [];
    readonly List<string> _codes = [];

    /// <summary>Distinct codes in index order - the order styles.xml must declare them in.</summary>
    public IReadOnlyList<string> Codes => _codes;

    public int GetOrAdd(string code)
    {
        if (string.IsNullOrEmpty(code))
            throw new InvalidOperationException(SR.FormatCodeEmpty);

        if (_indexes.TryGetValue(code, out var index))
            return index;

        index = _codes.Count;
        _indexes.Add(code, index);
        _codes.Add(code);
        return index;
    }
}
