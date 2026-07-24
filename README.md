# FakeExcelSerializer
Convert object to Excel file (.xlsx) [Open XML SpreadsheetML File Format]

## Getting Started
Supporting platform is .NET Standard 2.0, 2.1, .NET 5, .NET 6.

~~~
PM> Install-Package FakeExcelSerializer
~~~

## Usage
You can use `ExcelSerializer.ToFile` to create .xlsx file.

~~~
ExcelSerializer.ToFile(Users, "test.xlsx", ExcelSerializerOptions.Default);
~~~

## Notice

Folder creation permissions are required since a working folder will be used.

## Benchmark
N = 100 lines.

|              Method |   N |        Mean |     Error |    StdDev | Ratio |      Gen 0 |      Gen 1 |     Gen 2 |  Allocated |
|-------------------- |---- |------------:|----------:|----------:|------:|-----------:|-----------:|----------:|-----------:|
|           ClosedXml |   1 |    73.00 ms |  1.450 ms |  4.091 ms |  1.00 |          - |          - |         - |   5,738 KB |
| FakeExcelSerializer |   1 |    18.88 ms |  0.362 ms |  0.417 ms |  0.24 |          - |          - |         - |     126 KB |
|                     |     |             |           |           |       |            |            |           |            |
|           ClosedXml |  10 |   630.72 ms |  4.783 ms |  3.994 ms |  1.00 |  9000.0000 |  2000.0000 |         - |  52,663 KB |
| FakeExcelSerializer |  10 |    25.86 ms |  0.490 ms |  0.619 ms |  0.04 |   156.2500 |    31.2500 |         - |     661 KB |
|                     |     |             |           |           |       |            |            |           |            |
|           ClosedXml | 100 | 6,620.75 ms | 57.586 ms | 48.087 ms |  1.00 | 91000.0000 | 22000.0000 | 5000.0000 | 513,948 KB |
| FakeExcelSerializer | 100 |    77.63 ms |  1.491 ms |  1.395 ms |  0.01 |  1428.5714 |   142.8571 |         - |   6,005 KB |

## Example-1
If you pass an object, it will be converted to an Excel file.  
![image](https://user-images.githubusercontent.com/16958552/185727609-79b574e8-b40c-46dc-83c9-74b078a1f44a.png)
~~~
ExcelSerializer.ToFile(new string[] { "test", "test2" }, @"c:\test\test.xlsx", ExcelSerializerOptions.Default);
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

ExcelSerializer.ToFile(potals, @"c:\test\potals.xlsx", ExcelSerializerOptions.Default);
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

var newConfig = ExcelSerializerOptions.Default with
{
    HasHeaderRecord = true,
};
ExcelSerializer.ToFile(potals, @"c:\test\potalsEx.xlsx", newConfig);
~~~
## Example-4
Options can be set to display a title line and automatically adjust column widths.  
![image](https://user-images.githubusercontent.com/16958552/185727708-18201283-bb0b-46ba-a413-dbe34c20f3a3.png)
~~~
var newConfig = ExcelSerializerOptions.Default with
{
    CultureInfo = CultureInfo.InvariantCulture,
    HasHeaderRecord = true,
    HeaderTitles = new string[] { "Name", "Owner", "Level" },
    AutoFitColumns = true,
};
ExcelSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);
~~~

## Example-5
Optionally supports Autofilter.  

var newConfig = ExcelSerializerOptions.Default with
{
    CultureInfo = CultureInfo.InvariantCulture,
    HasHeaderRecord = true,
    HeaderTitles = new string[] { "Name", "Owner", "Level" },
    AutoFitlter = true,
};
ExcelSerializer.ToFile(potals, @"c:\test\potalsOp.xlsx", newConfig);

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
