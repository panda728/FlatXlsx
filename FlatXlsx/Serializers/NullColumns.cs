namespace FlatXlsx.Serializers;

/// <summary>A serializer whose null writes the same number of cells as its values do.</summary>
/// <remarks>
/// A nested object spans several columns, so writing one empty cell for a null instance shifts
/// every later value under the wrong heading. Serializers that know their span implement this;
/// for everything else a null is one empty cell, which is exact for scalars and unchanged
/// behaviour for custom serializers, whose span only their author knows.
/// </remarks>
internal interface INullColumnSpan
{
    void WriteNull(XlsxCellWriter writer, XlsxSerializerOptions options);
}

internal static class NullColumns
{
    // Cuts type cycles in the null walk: a TreeNode whose Parent is null must not recurse
    // TreeNode -> Parent -> TreeNode until MaxDepth. Thread-static because a writer is
    // single-threaded but distinct exports may run concurrently.
    [ThreadStatic]
    static HashSet<Type>? _inProgress;

    /// <summary>Writes the null representation of a member of type <typeparamref name="TMember"/>:
    /// the serializer's own span when it declares one, a single empty cell otherwise.
    /// Called from the compiled member expressions.</summary>
    public static void WriteFor<TMember>(XlsxCellWriter writer, XlsxSerializerOptions options)
    {
        if (options.GetRequiredSerializer<TMember>() is INullColumnSpan spanned)
            spanned.WriteNull(writer, options);
        else
            writer.WriteEmpty();
    }

    /// <summary>Runs a type's null walk once per branch: re-entering the same type collapses to
    /// one empty cell, so cyclic shapes terminate the way a null link terminates them at
    /// runtime.</summary>
    public static void Walk(Type type, XlsxCellWriter writer, XlsxSerializerOptions options,
        Action<XlsxCellWriter, XlsxSerializerOptions> writeNulls)
    {
        var guard = _inProgress ??= new HashSet<Type>();
        if (!guard.Add(type))
        {
            writer.WriteEmpty();
            return;
        }
        try
        {
            writeNulls(writer, options);
        }
        finally
        {
            guard.Remove(type);
        }
    }
}
