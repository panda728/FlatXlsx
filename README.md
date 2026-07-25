# FlatXlsx
Convert object to Excel file (.xlsx) [Open XML SpreadsheetML File Format]

FlatXlsx is the successor of [FakeExcelSerializer](https://github.com/panda728/FakeExcelSerializer), renamed and updated for .NET 10.

## Getting Started
Supported platforms: .NET Standard 2.0 / 2.1 (.NET Framework 4.6.2+, .NET 5/6/7), .NET 8, and .NET 10.

~~~
PM> Install-Package FlatXlsx
~~~

## Usage
You can use `XlsxSerializer.ToFile` to create .xlsx file. Options are optional; the defaults
just work.

~~~csharp
XlsxSerializer.ToFile(Users, "test.xlsx");
~~~

Every entry point is named after its destination: `ToFile`, `ToStream` (any `Stream`, seekable
or not), `ToBufferWriter` (any `IBufferWriter<byte>`), and `ToPipeWriterAsync`. Output is fully
streamed; no temporary files or working folder are used.

~~~csharp
// Stream
XlsxSerializer.ToStream(Users, stream);

// ASP.NET Core: write directly to the response without buffering a file
app.MapGet("/users.xlsx", (HttpResponse response) =>
{
    response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    XlsxSerializer.ToBufferWriter(GetUsers(), response.BodyWriter, XlsxSerializerOptions.Default);
    return Task.CompletedTask;
});
~~~

Async variants are available: `ToFileAsync`, `ToStreamAsync`, and `ToPipeWriterAsync`.
The PipeWriter overload flushes as data is produced, so pipe backpressure is honored.
`ToStreamAsync` never writes the stream synchronously (the zip container's few synchronous
tail writes are buffered and forwarded asynchronously), so an ASP.NET Core response body is a
valid destination for either overload.

An `IAsyncEnumerable<T>` source is accepted by the same three async entry points
(on .NET 8+, .NET 10 and .NET Standard 2.1), so a query's `AsAsyncEnumerable()` or a streaming
service response exports directly - rows are awaited as they arrive, never materialized first:

~~~csharp
await XlsxSerializer.ToFileAsync(db.Users.AsAsyncEnumerable(), "users.xlsx", cancellationToken: ct);
~~~

~~~csharp
await XlsxSerializer.ToFileAsync(Users, "test.xlsx", XlsxSerializerOptions.Default, cancellationToken);

// ASP.NET Core, fully async
app.MapGet("/users.xlsx", async (HttpResponse response, CancellationToken ct) =>
{
    response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    await XlsxSerializer.ToPipeWriterAsync(GetUsers(), response.BodyWriter, XlsxSerializerOptions.Default, ct);
});
~~~

## Supported types

| Category | Types | Available on |
|---|---|---|
| Numeric primitives | `bool` `byte` `sbyte` `short` `ushort` `int` `uint` `long` `ulong` `float` `double` `decimal` | all targets |
| Native-size integers | `nint` (`IntPtr`) / `nuint` (`UIntPtr`) | all targets |
| System.Numerics | `BigInteger` `Complex` | all targets |
| Text | `string` `char` | all targets |
| Common structs/classes | `Guid` `Enum` `DateTime` `DateTimeOffset` `TimeSpan` `Uri` `Version` | all targets |
| .NET Core 3.0+ types | `Rune` | .NET 8+ |
| .NET 5+ types | `Half` | .NET 8+ |
| .NET 6+ types | `DateOnly` `TimeOnly` | .NET 8+ |
| .NET 7+ types | `Int128` `UInt128` | .NET 8+ |

`Nullable<T>` of any supported struct, tuples, and object graphs (classes/records with public properties/fields) are also supported.
Values of `BigInteger` / `Int128` / `UInt128` exceeding 15 significant digits are stored in full in the cell,
but Excel displays numbers with at most 15 digits of precision.

## Out of scope

FlatXlsx converts one flat sequence of rows into one worksheet, fast. The following are out of
scope by design — if the export needs them, reach for a full spreadsheet library such as
ClosedXML instead of looking for a hidden option here:

- Multiple worksheets per workbook
- Styling beyond the built-in number/date formats, wrap and header freeze
  (fonts, colors, borders, merged cells)
- Formulas, charts, images, pivot tables
- Reading or editing existing workbooks (FlatXlsx is write-only)

## Notice

Output is streamed directly to the destination; no working folder is used.

- `SheetName` names the (single) worksheet; Excel's naming rules are enforced at write time,
  so an invalid name fails with a clear message instead of a workbook Excel repairs.
- An empty source writes nothing — unless `HeaderTitles` is set, in which case a header-only
  workbook is written so downstream consumers still receive a file.
- A null source throws `ArgumentNullException`; silence there would only hide a caller bug.

## Message languages

Exception messages follow `CultureInfo.CurrentUICulture`. English is built into the main
assembly; other languages ship as satellite assemblies (`ja/FlatXlsx.resources.dll`) that NuGet
restores alongside it.

- Nothing to configure: on a Japanese system the messages are Japanese.
- An untranslated language falls back to English, as does an application that trims or omits the
  satellite assemblies.
- This is separate from `XlsxSerializerOptions.CultureInfo`, which affects only the types whose
  text is culture-defined (`DateTimeOffset`, `Complex`) and itself defaults to the invariant
  culture. Numbers, dates and times are always stored invariantly, so the same code produces
  the same file on every machine.

To add a language, copy `FlatXlsx/Resources/Strings.resx` to `Strings.<culture>.resx`, translate
the `<value>` elements, and rebuild - the satellite assembly is produced automatically.

## Handling untrusted data

Values are frequently outside the caller's control, so the writer is built to keep any input
from producing a broken or dangerous workbook:

| Concern | Behaviour |
|---|---|
| Markup in values (`<`, `&`, `</t></is></c>`) | Escaped as XML text; it can never become elements. Header titles included. |
| Control characters (`\0` and other C0 codes) | Dropped. XML 1.0 forbids them even escaped, and Excel rejects a file that contains them. |
| Formula injection (`=cmd\|...`) | Not applicable: values are written as string or number cells, never as `<f>` formulas, so Excel does not evaluate them. Re-exporting the sheet to CSV is a separate matter. |
| Circular references / deep nesting | `MaxDepth` (default 64) aborts with a clear exception. |
| Collections expanding across columns | Excel's 16,384-column limit is enforced with an exception instead of silently emitting a corrupt file. Row (1,048,576) and cell-length (32,767) limits likewise. |
| High-cardinality strings | The shared-string table is capped by `MaxSharedStrings` (default 1,000,000); beyond it, values are written as inline strings, so memory stays bounded. |

The row source is enumerated exactly once, so a query or forward-only reader can be passed
directly. With `AutoFitColumns` the first `AutoFitSampleRows` rows (default 200) are buffered to
measure widths — bounded, never the whole sequence.

`XlsxCellWriter` instances are not thread-safe; serializer providers and their caches are.

## Benchmark

FlatXlsx 1.0.0 vs ClosedXML 0.105.0, .NET 10 (BenchmarkDotNet 0.15.8, ShortRun; N = 100 lines).
Three variants are measured so the comparison is fair to ClosedXML:

- **ClosedXmlNaive** — beginner-style code: cell-by-cell writes, per-cell number formats,
  `ColumnsUsed().AdjustToContents()` column sizing.
- **ClosedXmlOptimized** — tuned code: bulk `InsertData`, column-level number formats set once,
  no `AdjustToContents` (its per-cell measurement is ClosedXML's slowest feature).
- **FlatXlsx** — header row + approximate column auto-fit enabled (`AutoFitColumns = true`),
  so its feature set is comparable to the naive variant.

| Method             | N   | Mean        | Ratio | Allocated    | Alloc Ratio |
|------------------- |---- |------------:|------:|-------------:|------------:|
| ClosedXmlNaive     | 1   |    42.52 ms |  1.00 |   5,281.3 KB |       1.000 |
| ClosedXmlOptimized | 1   |     8.74 ms |  0.21 |   1,238.1 KB |       0.234 |
| FlatXlsx           | 1   |     0.80 ms |  0.02 |      23.7 KB |       0.004 |
| ClosedXmlNaive     | 10  |   167.99 ms |  1.00 |  48,535.6 KB |       1.000 |
| ClosedXmlOptimized | 10  |    30.29 ms |  0.18 |   9,314.8 KB |       0.192 |
| FlatXlsx           | 10  |     2.21 ms |  0.01 |      25.7 KB |       0.001 |
| ClosedXmlNaive     | 100 | 1,605.45 ms | 1.000 | 471,633.7 KB |       1.000 |
| ClosedXmlOptimized | 100 |   332.12 ms | 0.207 |  82,651.5 KB |       0.175 |
| FlatXlsx           | 100 |    15.18 ms | 0.009 |      25.7 KB |       0.000 |

Even against well-tuned ClosedXML, FlatXlsx is ~20x faster; output is fully streamed,
so allocations stay flat (~26 KB) regardless of row count.
For large data sets, `XlsxSerializerOptions.CompressionLevel = CompressionLevel.Fastest`
trades a slightly larger file for even faster serialization.

## Examples

Every example below is runnable:

~~~bash
dotnet run --project FlatXlsx.Console
~~~

writes each one to an `output/` folder — one workbook per example, plus a kitchen-sink
workbook that exercises every supported type at once.

## Example-1
If you pass an object, it will be converted to an Excel file.  
![image](https://user-images.githubusercontent.com/16958552/185727609-79b574e8-b40c-46dc-83c9-74b078a1f44a.png)
~~~csharp
XlsxSerializer.ToFile(new string[] { "test", "test2" }, @"c:\test\test.xlsx");
~~~

## Example-2
Passing a class expands the property into a column.  
![image](https://user-images.githubusercontent.com/16958552/185727657-3e41dea7-1af4-4a52-99bd-1457f895b564.png)
~~~csharp
public class Portal
{
    public string Name { get; set; }
    public string Owner { get; set; }
    public int Level { get; set; }
}

var potals = new Portal[] {
    new Portal { Name = "Portal1", Owner = "panda728", Level = 8 },
    new Portal { Name = "Portal2", Owner = "panda728", Level = 1 },
    new Portal { Name = "Portal3", Owner = "panda728", Level = 2 },
};

XlsxSerializer.ToFile(potals, @"c:\test\potals.xlsx", XlsxSerializerOptions.Default);
~~~
## Example-3
To add a header row with your own titles, set `HeaderTitles` — that one setting is the whole ask.
To use the member names as titles instead, set `HasHeaderRow = true`.
![image](https://user-images.githubusercontent.com/16958552/187447183-1c0af135-8407-4c79-be8d-0b4875973a79.png)
~~~csharp
XlsxSerializer.ToFile(potals, @"c:\test\potals.xlsx",
    new XlsxSerializerOptions { HeaderTitles = new[] { "Name", "Owner", "Level" } });
~~~

Column titles and order can also live on the class itself, next to the members they describe.
`[DataMember]` is honored too, for types already annotated for other serializers.
~~~csharp
public class Portal
{
    [XlsxColumn("Name Ex", Order = 3)]
    public string Name { get; set; }
    [XlsxColumn("Owner Ex", Order = 1)]
    public string Owner { get; set; }
    [XlsxColumn("Level Ex", Order = 2)]
    public int Level { get; set; }
    [XlsxIgnore]
    public string InternalNote { get; set; }
}

XlsxSerializer.ToFile(potals, @"c:\test\potalsEx.xlsx",
    new XlsxSerializerOptions { HasHeaderRow = true });
~~~
## Example-4
Options can be set to display a title line and automatically adjust column widths.  
![image](https://user-images.githubusercontent.com/16958552/185727708-18201283-bb0b-46ba-a413-dbe34c20f3a3.png)
~~~csharp
var newConfig = XlsxSerializerOptions.Default with
{
    HeaderTitles = new[] { "Name", "Owner", "Level" },
    AutoFitColumns = true,
};
XlsxSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);
~~~

## Example-5
Optionally supports Autofilter.  

~~~csharp
var newConfig = XlsxSerializerOptions.Default with
{
    HeaderTitles = new[] { "Name", "Owner", "Level" },
    AutoFilter = true,
};
XlsxSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);
~~~

## Example-6
To change how one type is written, set `CustomSerializers` — matched by value type, consulted
before the built-in serializers, no provider wiring needed. The shipped
`EnumNumberXlsxSerializer<T>` (numeric enums) plugs in the same way.

~~~csharp
class YesNoSerializer : IXlsxSerializer<bool>
{
    public void WriteTitle(XlsxCellWriter writer, bool value, XlsxSerializerOptions options, string name = "value")
        => writer.Write(name);
    public void Serialize(XlsxCellWriter writer, bool value, XlsxSerializerOptions options)
        => writer.Write(value ? "YES" : "NO");
}

XlsxSerializer.ToFile(rows, @"c:\test\yesno.xlsx",
    new XlsxSerializerOptions { CustomSerializers = new IXlsxSerializer[] { new YesNoSerializer() } });
~~~

## Example-7
One column can have its own number format without touching the sheet-wide ones: register the
format code in `CustomFormats` and refer to it by index from a serializer. The display format
changes; the stored value stays a real number.

~~~csharp
class PercentSerializer : IXlsxSerializer<double>
{
    public void WriteTitle(XlsxCellWriter writer, double value, XlsxSerializerOptions options, string name = "value")
        => writer.Write(name);
    public void Serialize(XlsxCellWriter writer, double value, XlsxSerializerOptions options)
        => writer.Write(value, customFormat: 0);   // -> CustomFormats[0]
}

XlsxSerializer.ToFile(rows, @"c:\test\percent.xlsx", new XlsxSerializerOptions
{
    CustomFormats = new[] { "0.0%" },
    CustomSerializers = new IXlsxSerializer[] { new PercentSerializer() },
});
~~~

## Acknowledgements

FlatXlsx's serialization pipeline is a direct descendant of
[Cysharp/WebSerializer](https://github.com/Cysharp/WebSerializer).
The original FakeExcelSerializer ported WebSerializer's architecture —
provider-based serializer resolution and compiled member accessors over `IEnumerable<T>` —
and swapped the output layer to SpreadsheetML. If you need the same zero-allocation
approach for query strings or form data on the web side, check out WebSerializer itself.

## License
This library is licensed under the MIT License.
