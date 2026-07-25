using System.Collections.ObjectModel;
using System.Numerics;
using System.Runtime.Serialization;

namespace FlatXlsx.Tests.Support;

/// <summary>
/// One supported type, plus what a reader must see after a value of that type is written.
/// </summary>
/// <remarks>
/// The contract belongs to <see cref="IXlsxSerializer{T}"/>, not to any one implementation, so
/// the clauses in SerializerContractTests are applied to every entry in <see cref="All"/>.
/// Adding a type here subjects it to the whole contract at once.
/// </remarks>
abstract class SerializerCase
{
    public abstract byte[] Write(XlsxSerializerOptions options);

    /// <summary>What the cells of the data row must read back as. null means an empty cell.</summary>
    public abstract string?[] ExpectedCells { get; }

    public static SerializerCase Of<T>(T value, params string?[] expected) => new Case<T>(value, expected);

    sealed class Case<T>(T value, string?[] expected) : SerializerCase
    {
        public override byte[] Write(XlsxSerializerOptions options) => Xlsx.Write(new[] { value }, options);
        public override string?[] ExpectedCells => expected;
    }

    public enum Fruit { Apple, Orange }

    public sealed class Portal
    {
        [DataMember(Name = "Owner Ex", Order = 1)]
        public string? Owner { get; set; }
        [DataMember(Name = "Name Ex", Order = 2)]
        public string Name { get; set; } = "";
        [DataMember(Name = "Level Ex", Order = 3)]
        public int Level { get; set; }
    }

    static readonly CultureInfoDependent _cultureDependent = new();

    /// <summary>Values whose text is defined as "formatted with options.CultureInfo".</summary>
    sealed class CultureInfoDependent
    {
        public readonly DateTimeOffset Offset = new(2000, 1, 1, 10, 30, 0, TimeSpan.FromHours(9));
        public readonly Complex Complex = new(1, 2);
    }

    public static readonly IReadOnlyDictionary<string, SerializerCase> All = new Dictionary<string, SerializerCase>
    {
        // Numbers keep their exact value; a spreadsheet that rounds or reformats them is wrong.
        ["bool"] = Of(true, "True"),
        ["byte"] = Of((byte)200, "200"),
        ["sbyte"] = Of((sbyte)-100, "-100"),
        ["short"] = Of(short.MinValue, "-32768"),
        ["ushort"] = Of(ushort.MaxValue, "65535"),
        ["int"] = Of(int.MinValue, "-2147483648"),
        ["uint"] = Of(uint.MaxValue, "4294967295"),
        ["long"] = Of(long.MinValue, "-9223372036854775808"),
        ["ulong"] = Of(ulong.MaxValue, "18446744073709551615"),
        ["float"] = Of(1.5f, "1.5"),
        ["double"] = Of(-0.25d, "-0.25"),
        ["decimal"] = Of(decimal.MinValue, "-79228162514264337593543950335"),
        ["Half"] = Of((Half)1.5, "1.5"),
        ["Int128"] = Of(Int128.MaxValue, "170141183460469231731687303715884105727"),
        ["UInt128"] = Of(UInt128.MaxValue, "340282366920938463463374607431768211455"),
        ["BigInteger"] = Of(BigInteger.Parse("123456789012345678901234567890"), "123456789012345678901234567890"),
        ["nint"] = Of((nint)(-777), "-777"),
        ["nuint"] = Of((nuint)12345, "12345"),

        // Text arrives verbatim.
        ["string"] = Of("hello", "hello"),
        ["char"] = Of('A', "A"),
        ["Rune"] = Of(new System.Text.Rune(0x1F600), "\U0001F600"),
        ["Guid"] = Of(Guid.Parse("2f1c8f0e-0000-4000-8000-000000000001"), "2f1c8f0e-0000-4000-8000-000000000001"),
        ["Enum"] = Of(Fruit.Orange, "Orange"),
        ["Uri"] = Of(new Uri("http://example.com/portals"), "http://example.com/portals"),
        ["Version"] = Of(new Version(1, 2, 3, 4), "1.2.3.4"),
        ["TimeSpan"] = Of(TimeSpan.FromHours(10), "10:00:00"),

        // Dates are stored as ISO 8601 so that the spreadsheet reads them as dates, not text.
        ["DateTime"] = Of(new DateTime(2023, 1, 2, 3, 4, 5), "2023-01-02T03:04:05"),
        ["DateTime(midnight)"] = Of(new DateTime(2023, 1, 2), "2023-01-02T00:00:00"),
        ["DateOnly"] = Of(new DateOnly(2023, 1, 2), "2023-01-02T00:00:00"),
        ["TimeOnly"] = Of(new TimeOnly(3, 4, 5), "1900-01-01T03:04:05"),
        ["DateTimeOffset"] = Of(_cultureDependent.Offset, _cultureDependent.Offset.ToString(XlsxSerializerOptions.Default.CultureInfo)),
        ["Complex"] = Of(_cultureDependent.Complex, _cultureDependent.Complex.ToString(XlsxSerializerOptions.Default.CultureInfo)),

        // Absent values leave the cell empty rather than writing a placeholder.
        ["int?(null)"] = Of((int?)null, new string?[] { null }),
        ["int?(value)"] = Of((int?)42, "42"),
        ["string(null)"] = Of((string?)null, new string?[] { null }),
        ["string(empty)"] = Of("", new string?[] { null }),
        ["Uri(null)"] = Of((Uri?)null, new string?[] { null }),

        // Negative magnitudes take a different path through the formatters than positive ones.
        ["BigInteger(negative)"] = Of(BigInteger.Parse("-123456789012345678901234567890"), "-123456789012345678901234567890"),
        ["Int128(negative)"] = Of(Int128.MinValue, "-170141183460469231731687303715884105728"),
        ["Half?(null)"] = Of((Half?)null, new string?[] { null }),
        ["BigInteger?(null)"] = Of((BigInteger?)null, new string?[] { null }),

        // Every tuple arity is served by its own generated serializer, so each one owes the
        // contract separately.
        ["Tuple(1)"] = Of(Tuple.Create(1), "1"),
        ["Tuple(2)"] = Of(Tuple.Create(1, 2), "1", "2"),
        ["Tuple(3)"] = Of(Tuple.Create(1, 2, 3), "1", "2", "3"),
        ["Tuple(4)"] = Of(Tuple.Create(1, 2, 3, 4), "1", "2", "3", "4"),
        ["Tuple(5)"] = Of(Tuple.Create(1, 2, 3, 4, 5), "1", "2", "3", "4", "5"),
        ["Tuple(6)"] = Of(Tuple.Create(1, 2, 3, 4, 5, 6), "1", "2", "3", "4", "5", "6"),
        ["Tuple(7)"] = Of(Tuple.Create(1, 2, 3, 4, 5, 6, 7), "1", "2", "3", "4", "5", "6", "7"),
        ["Tuple(8)"] = Of(Tuple.Create(1, 2, 3, 4, 5, 6, 7, 8), "1", "2", "3", "4", "5", "6", "7", "8"),
        ["ValueTuple(1)"] = Of(ValueTuple.Create(1), "1"),
        ["ValueTuple(2)"] = Of((1, 2), "1", "2"),
        ["ValueTuple(3)"] = Of((1, 2, 3), "1", "2", "3"),
        ["ValueTuple(4)"] = Of((1, 2, 3, 4), "1", "2", "3", "4"),
        ["ValueTuple(5)"] = Of((1, 2, 3, 4, 5), "1", "2", "3", "4", "5"),
        ["ValueTuple(6)"] = Of((1, 2, 3, 4, 5, 6), "1", "2", "3", "4", "5", "6"),
        ["ValueTuple(7)"] = Of((1, 2, 3, 4, 5, 6, 7), "1", "2", "3", "4", "5", "6", "7"),
        ["ValueTuple(8)"] = Of((1, 2, 3, 4, 5, 6, 7, 8), "1", "2", "3", "4", "5", "6", "7", "8"),

        // Composite values spread across consecutive cells.
        ["ValueTuple(mixed)"] = Of((1, "two"), "1", "two"),
        ["Tuple(mixed)"] = Of(Tuple.Create(1, "two"), "1", "two"),
        ["Collection"] = Of(new Collection<string> { "a", "b" }, "a", "b"),
        ["Dictionary"] = Of(new Dictionary<string, int> { ["k1"] = 1, ["k2"] = 2 }, "k1", "1", "k2", "2"),
        ["KeyValuePair"] = Of(new KeyValuePair<string, int>("k1", 1), "k1", "1"),
        ["KeyValuePair(null key)"] = Of(new KeyValuePair<string?, int>(null, 9), null, "9"),
        ["object(fallback)"] = Of((object)"boxed", "boxed"),
        ["object graph"] = Of(new Portal { Owner = null, Name = "Portal1", Level = 8 }, null, "Portal1", "8"),
    };

    public static TheoryData<string> Names
    {
        get
        {
            var data = new TheoryData<string>();
            foreach (var name in All.Keys)
                data.Add(name);
            return data;
        }
    }
}
