# FlatXlsx
Convert object to Excel file (.xlsx) [Open XML SpreadsheetML File Format]

FlatXlsx is the successor of [FakeExcelSerializer](https://github.com/panda728/FakeExcelSerializer), renamed and updated for .NET 10.

## Getting Started
Supported platforms: .NET Standard 2.0 / 2.1 (.NET Framework 4.6.2+, .NET 5/6/7), .NET 8, and .NET 10.

~~~
PM> Install-Package FlatXlsx
~~~

## Usage
You can use `XlsxSerializer.ToFile` to create .xlsx file.

~~~csharp
XlsxSerializer.ToFile(Users, "test.xlsx", XlsxSerializerOptions.Default);
~~~

`ToStream` writes to any `Stream` (it does not need to be seekable), and `To` writes to any
`IBufferWriter<byte>` — including `System.IO.Pipelines.PipeWriter`. Output is fully streamed;
no temporary files or working folder are used.

~~~csharp
// Stream
XlsxSerializer.ToStream(Users, stream, XlsxSerializerOptions.Default);

// ASP.NET Core: write directly to the response without buffering a file
app.MapGet("/users.xlsx", (HttpResponse response) =>
{
    response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    XlsxSerializer.To(GetUsers(), response.BodyWriter, XlsxSerializerOptions.Default);
    return Task.CompletedTask;
});
~~~

Async variants are available: `ToFileAsync`, `ToStreamAsync`, and `ToAsync(PipeWriter)`.
The PipeWriter overload flushes as data is produced, so pipe backpressure is honored.

~~~csharp
await XlsxSerializer.ToFileAsync(Users, "test.xlsx", XlsxSerializerOptions.Default, cancellationToken);

// ASP.NET Core, fully async
app.MapGet("/users.xlsx", async (HttpResponse response, CancellationToken ct) =>
{
    response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    await XlsxSerializer.ToAsync(GetUsers(), response.BodyWriter, XlsxSerializerOptions.Default, ct);
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

## Notice

Output is streamed directly to the destination; no working folder is used.
(`XlsxSerializerOptions.WorkPath` is obsolete and ignored.)

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
directly. With `AutoFitColumns` the first `AutoFitDepth` rows (default 200) are buffered to
measure widths — bounded, never the whole sequence.

`XlsxWriter` instances are not thread-safe; serializer providers and their caches are.

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
| ClosedXmlNaive     | 1   |    26.19 ms |  1.00 |   5,281.1 KB |       1.000 |
| ClosedXmlOptimized | 1   |     8.26 ms |  0.32 |   1,239.0 KB |       0.235 |
| FlatXlsx           | 1   |     0.70 ms |  0.03 |      21.5 KB |       0.004 |
| ClosedXmlNaive     | 10  |   155.48 ms |  1.00 |  48,536.1 KB |       1.000 |
| ClosedXmlOptimized | 10  |    30.86 ms |  0.20 |   9,312.5 KB |       0.192 |
| FlatXlsx           | 10  |     2.09 ms |  0.01 |      21.5 KB |       0.000 |
| ClosedXmlNaive     | 100 | 1,904.67 ms | 1.000 | 471,629.7 KB |       1.000 |
| ClosedXmlOptimized | 100 |   297.29 ms | 0.156 |  82,638.1 KB |       0.175 |
| FlatXlsx           | 100 |    16.15 ms | 0.008 |      21.5 KB |       0.000 |

Even against well-tuned ClosedXML, FlatXlsx is ~18x faster; output is fully streamed,
so allocations stay flat (~21 KB) regardless of row count.
For large data sets, `XlsxSerializerOptions.CompressionLevel = CompressionLevel.Fastest`
trades a slightly larger file for even faster serialization.

## Example-1
If you pass an object, it will be converted to an Excel file.  
![image](https://user-images.githubusercontent.com/16958552/185727609-79b574e8-b40c-46dc-83c9-74b078a1f44a.png)
~~~csharp
XlsxSerializer.ToFile(new string[] { "test", "test2" }, @"c:\test\test.xlsx", XlsxSerializerOptions.Default);
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
By setting attributes on the class, you can specify the name of the title or change the order of the columns.  
![image](https://user-images.githubusercontent.com/16958552/187447183-1c0af135-8407-4c79-be8d-0b4875973a79.png)
~~~csharp
public class Portal
{
    [DataMember(Name = "Name Ex", Order = 3)]
    public string Name { get; set; }
    [DataMember(Name = "Owner Ex", Order = 1)]
    public string Owner { get; set; }
    [DataMember(Name = "Level Ex", Order = 2)]
    public int Level { get; set; }
}

var potals = new Portal[] {
    new Portal { Name = "Portal1", Owner = "panda728", Level = 8 },
    new Portal { Name = "Portal2", Owner = "panda728", Level = 1 },
    new Portal { Name = "Portal3", Owner = "panda728", Level = 2 },
};

var newConfig = XlsxSerializerOptions.Default with
{
    HasHeaderRecord = true,
};
XlsxSerializer.ToFile(potals, @"c:\test\potalsEx.xlsx", newConfig);
~~~
## Example-4
Options can be set to display a title line and automatically adjust column widths.  
![image](https://user-images.githubusercontent.com/16958552/185727708-18201283-bb0b-46ba-a413-dbe34c20f3a3.png)
~~~csharp
var newConfig = XlsxSerializerOptions.Default with
{
    CultureInfo = CultureInfo.InvariantCulture,
    HasHeaderRecord = true,
    HeaderTitles = new string[] { "Name", "Owner", "Level" },
    AutoFitColumns = true,
};
XlsxSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);
~~~

## Example-5
Optionally supports Autofilter.  

~~~csharp
var newConfig = XlsxSerializerOptions.Default with
{
    CultureInfo = CultureInfo.InvariantCulture,
    HasHeaderRecord = true,
    HeaderTitles = new string[] { "Name", "Owner", "Level" },
    AutoFilter = true,
};
XlsxSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);
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
