namespace FlatXlsx;

/// <summary>
/// Interns each distinct cell string once and hands out the index the cell refers to it by.
/// </summary>
/// <remarks>
/// Owns the whole shared-string policy: index assignment, the memory cap, and the write order
/// of strings.xml. Entries are kept in an explicit list rather than relying on dictionary
/// enumeration order, so the file's order is guaranteed by construction, not by an
/// implementation detail of Dictionary.
/// </remarks>
internal sealed class SharedStringTable(int maxCount)
{
    readonly Dictionary<string, int> _indexes = new();
    readonly List<string> _values = new();

    public int Count => _values.Count;

    /// <summary>Entries in index order - the order strings.xml must list them in.</summary>
    public IReadOnlyList<string> Values => _values;

    /// <summary>Returns the value's index, adding it if there is room.
    /// False means the table is full and the value is not in it - the caller writes the value
    /// inline instead, so a full table costs file size, never correctness.</summary>
    public bool TryGetOrAdd(string value, out int index)
    {
        if (_indexes.TryGetValue(value, out index))
            return true;

        if (_values.Count >= maxCount)
        {
            index = -1;
            return false;
        }

        index = _values.Count;
        _indexes.Add(value, index);
        _values.Add(value);
        return true;
    }
}
