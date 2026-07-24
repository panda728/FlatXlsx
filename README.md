# FlatXlsx
Convert object to Excel file (.xlsx) [Open XML SpreadsheetML File Format]

FlatXlsx is the successor of [FakeXlsxSerializer](https://github.com/panda728/FakeXlsxSerializer), renamed and updated for .NET 10.

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

## Benchmark

FlatXlsx 1.0.0 vs ClosedXML 0.105.0, .NET 10 (BenchmarkDotNet 0.15.8, ShortRun; N = 100 lines):

| Method    | N   | Mean        | Ratio | Gen0       | Gen1      | Gen2      | Allocated    | Alloc Ratio |
|---------- |---- |------------:|------:|-----------:|----------:|----------:|-------------:|------------:|
| ClosedXml | 1   |    49.78 ms |  1.00 |   600.0000 |  200.0000 |         - |   5,269.3 KB |       1.000 |
| FlatXlsx  | 1   |     0.75 ms |  0.02 |     1.9531 |         - |         - |      17.2 KB |       0.003 |
| ClosedXml | 10  |   172.17 ms |  1.00 |  5333.3333 |  333.3333 |         - |  48,481.2 KB |       1.000 |
| FlatXlsx  | 10  |     2.85 ms |  0.02 |          - |         - |         - |      17.2 KB |       0.000 |
| ClosedXml | 100 | 1,781.70 ms |  1.00 | 53000.0000 | 9000.0000 | 2000.0000 | 471,086.1 KB |       1.000 |
| FlatXlsx  | 100 |    22.01 ms |  0.01 |          - |         - |         - |      17.2 KB |       0.000 |

Output is fully streamed, so allocations stay flat (~17 KB) regardless of row count.

## Example-1
If you pass an object, it will be converted to an Excel file.  
![image](https://user-images.githubusercontent.com/16958552/185727609-79b574e8-b40c-46dc-83c9-74b078a1f44a.png)
~~~
XlsxSerializer.ToFile(new string[] { "test", "test2" }, @"c:\test\test.xlsx", XlsxSerializerOptions.Default);
~~~

## Example-2
Passing a class expands the property into a column.  
![image](https://user-images.githubusercontent.com/16958552/185727657-3e41dea7-1af4-4a52-99bd-1457f895b564.png)
~~~
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
~~~
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
~~~
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

var newConfig = XlsxSerializerOptions.Default with
{
    CultureInfo = CultureInfo.InvariantCulture,
    HasHeaderRecord = true,
    HeaderTitles = new string[] { "Name", "Owner", "Level" },
    AutoFitlter = true,
};
XlsxSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);

## Note
For the method of retrieving values from IEnumerable\<T\>, Cysharp's WebSerializer method is used.

　https://github.com/Cysharp/WebSerializer
  
The following page provides information on how to return to OpenOfficeXml.

　https://gist.github.com/iso2022jp/721df3095f4df512bfe2327503ea1119

　https://docs.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/2c5dee00-eff2-4b22-92b6-0738acd4475e
 
## Extensions Sample

WindowsForm's DataGridView to .xlsx

https://github.com/panda728/DataGridViewDump

## Link
CSV File output version
　https://github.com/panda728/FakeCsvSerializer

## License
This library is licensed under the MIT License.
