// Runs every example from README.md, one output file per example, so the samples
// can be watched working instead of taken on faith. Finishes with a kitchen-sink
// workbook that exercises every supported type at once.
using Bogus;
using FlatXlsx;
using System.Diagnostics;
using System.Globalization;
using static Bogus.DataSets.Name;

var outDir = Path.Combine(Environment.CurrentDirectory, "output");
if (Directory.Exists(outDir))
    Directory.Delete(outDir, recursive: true);
Directory.CreateDirectory(outDir);

string Out(string name) => Path.Combine(outDir, name);
void Report(string section, string what, string file)
    => Console.WriteLine($"[{section,-9}] {what,-58} -> {Path.GetFileName(file)}");

var portals = new Portal[] {
    new Portal { Name = "Portal1", Owner = "panda728", Level = 8 },
    new Portal { Name = "Portal2", Owner = "panda728", Level = 1 },
    new Portal { Name = "Portal3", Owner = "panda728", Level = 2 },
};

// ---- Usage: the smallest possible call -------------------------------------------------
XlsxSerializer.ToFile(portals, Out("usage.xlsx"));
Report("Usage", "One line, no options", Out("usage.xlsx"));

// ---- Usage: ToStream (any Stream; it does not need to be seekable) ---------------------
using (var ms = new MemoryStream())
{
    XlsxSerializer.ToStream(portals, ms);
    Report("Usage", $"ToStream wrote {ms.Length:#,##0} bytes into a MemoryStream", "(no file)");
}

// ---- Usage: async variant --------------------------------------------------------------
await XlsxSerializer.ToFileAsync(portals, Out("usage-async.xlsx"));
Report("Usage", "ToFileAsync (ToStreamAsync never writes synchronously)", Out("usage-async.xlsx"));

// ---- Usage: an IAsyncEnumerable source streams straight in ------------------------------
static async IAsyncEnumerable<Portal> PortalsAsync()
{
    for (var i = 1; i <= 3; i++)
    {
        await Task.Delay(1);   // rows arriving from a query or a service
        yield return new Portal { Name = $"Portal{i}", Owner = "panda728", Level = i };
    }
}
await XlsxSerializer.ToFileAsync(PortalsAsync(), Out("usage-asyncsource.xlsx"));
Report("Usage", "IAsyncEnumerable source, rows awaited as they arrive", Out("usage-asyncsource.xlsx"));

// ---- Example-1: any sequence becomes a sheet -------------------------------------------
XlsxSerializer.ToFile(new string[] { "test", "test2" }, Out("example1.xlsx"));
Report("Example-1", "A string array, one value per row", Out("example1.xlsx"));

// ---- Example-2: a class expands into columns -------------------------------------------
XlsxSerializer.ToFile(portals, Out("example2.xlsx"));
Report("Example-2", "Properties become columns", Out("example2.xlsx"));

// ---- Example-3: header titles are one setting ------------------------------------------
XlsxSerializer.ToFile(portals, Out("example3-titles.xlsx"),
    new XlsxSerializerOptions { HeaderTitles = new[] { "Name", "Owner", "Level" } });
Report("Example-3", "HeaderTitles alone asks for the header row", Out("example3-titles.xlsx"));

// ---- Example-3: titles and order can live on the class ---------------------------------
var annotated = new AnnotatedPortal[] {
    new AnnotatedPortal { Name = "Portal1", Owner = "panda728", Level = 8, InternalNote = "secret" },
    new AnnotatedPortal { Name = "Portal2", Owner = "panda728", Level = 1, InternalNote = "secret" },
};
XlsxSerializer.ToFile(annotated, Out("example3-attributes.xlsx"),
    new XlsxSerializerOptions { HasHeaderRow = true });
Report("Example-3", "[XlsxColumn] renames/orders, [XlsxIgnore] excludes", Out("example3-attributes.xlsx"));

// ---- Example-4: auto-fit column widths -------------------------------------------------
XlsxSerializer.ToFile(portals, Out("example4.xlsx"), new XlsxSerializerOptions
{
    HeaderTitles = new[] { "Name", "Owner", "Level" },
    AutoFitColumns = true,
});
Report("Example-4", "AutoFitColumns sizes columns to their content", Out("example4.xlsx"));

// ---- Example-5: auto-filter ------------------------------------------------------------
XlsxSerializer.ToFile(portals, Out("example5.xlsx"), new XlsxSerializerOptions
{
    HeaderTitles = new[] { "Name", "Owner", "Level" },
    AutoFilter = true,
});
Report("Example-5", "AutoFilter covers the header and every row", Out("example5.xlsx"));

// ---- Example-6: a custom serializer is one setting -------------------------------------
XlsxSerializer.ToFile(
    new[] { (Name: "Portal1", Active: true), (Name: "Portal2", Active: false) },
    Out("example6.xlsx"),
    new XlsxSerializerOptions { CustomSerializers = new IXlsxSerializer[] { new YesNoSerializer() } });
Report("Example-6", "CustomSerializers: bool cells become YES/NO", Out("example6.xlsx"));

// ---- Example-7: one column with its own number format ----------------------------------
XlsxSerializer.ToFile(
    new[] { (Item: "CPU", Load: 0.125), (Item: "RAM", Load: 0.5) },
    Out("example7.xlsx"),
    new XlsxSerializerOptions
    {
        CustomFormats = new[] { "0.0%" },
        CustomSerializers = new IXlsxSerializer[] { new PercentSerializer() },
    });
Report("Example-7", "CustomFormats: a percent column, sheet formats untouched", Out("example7.xlsx"));

// ---- Kitchen sink: every supported type in one workbook --------------------------------
Randomizer.Seed = new Random(8675309);

var fruit = new[] { "apple", "banana", "orange", "strawberry", "kiwi" };

var orderIds = 0;
var testOrders = new Faker<Order>()
    .StrictMode(true)
    .RuleFor(o => o.OrderId, f => orderIds++)
    .RuleFor(o => o.Item, f => f.PickRandom(fruit))
    .RuleFor(o => o.Quantity, f => f.Random.Number(-10, 10))
    .RuleFor(o => o.LotNumber, f => f.Random.Int(0, 100).OrNull(f, .8f));

var userIds = 0;
var testUsers = new Faker<User>()
    .CustomInstantiator(f => new User(userIds++, f.Random.Replace("###-##-####")))
    .RuleFor(u => u.Gender, f => f.PickRandom<Gender>())
    .RuleFor(u => u.FirstName, (f, u) => f.Name.FirstName(u.Gender))
    .RuleFor(u => u.LastName, (f, u) => f.Name.LastName(u.Gender))
    .RuleFor(u => u.Avatar, f => f.Internet.Avatar())
    .RuleFor(u => u.UserName, (f, u) => f.Internet.UserName(u.FirstName, u.LastName))
    .RuleFor(u => u.Email, (f, u) => f.Internet.Email(u.FirstName, u.LastName))
    .RuleFor(u => u.SomethingUnique, f => $"Value {f.UniqueIndex}")
    .RuleFor(u => u.TimeStamp, f => f.Date.Recent())
    .RuleFor(u => u.CreateTime, f => f.Date.Recent())
    .RuleFor(u => u.DateOnlyValue, f => f.Date.RecentDateOnly())
    .RuleFor(u => u.TimeOnlyValue, f => f.Date.RecentTimeOnly())
    .RuleFor(u => u.TimeSpanValue, f => f.Date.Recent() - f.Date.Past())
    .RuleFor(u => u.DateTimeOffsetValue, f => f.Date.Recent())
    .RuleFor(u => u.Fallback, (f, u) => (object)userIds)
    .RuleFor(u => u.Uri, f => new Uri(f.Internet.Url()))
    .RuleFor(u => u.SomeGuid, f => f.Random.Guid())
    .RuleFor(u => u.SendFlag, f => userIds % 3 == 0)
    .RuleFor(u => u.CartId, f => f.Random.Guid())
    .RuleFor(u => u.FullName, (f, u) => u.FirstName + " " + u.LastName)
    .RuleFor(u => u.Orders, f => testOrders.Generate(3).ToList())
    .RuleFor(o => o.DoubleValue, f => f.Random.Double(-1000, 1000))
    .RuleFor(o => o.Char, f => (char)f.Random.Int(65, 65 + 26))
    .RuleFor(o => o.Escape, f => "</>\"'&");

var users = testUsers.Generate(10);

var sw = Stopwatch.StartNew();
XlsxSerializer.ToFile(users, Out("kitchen-sink.xlsx"), new XlsxSerializerOptions
{
    CultureInfo = CultureInfo.CurrentCulture,   // opt-in: localize DateTimeOffset text
    SheetName = "KitchenSink",
    HasHeaderRow = true,
    AutoFitColumns = true,
    AutoFilter = true,
});
sw.Stop();
Report("Sink", $"{users.Count} rows, every supported type, {sw.ElapsedMilliseconds:#,##0}ms", Out("kitchen-sink.xlsx"));

Console.WriteLine();
Console.WriteLine($"All workbooks are in: {outDir}");

if (!Console.IsInputRedirected)
{
    Console.WriteLine("press Enter to close...");
    Console.ReadLine();
}

/// <summary>README Example-2: plain properties become columns.</summary>
public class Portal
{
    public string Name { get; set; } = "";
    public string Owner { get; set; } = "";
    public int Level { get; set; }
}

/// <summary>README Example-3 declares this as Portal; renamed here so both variants
/// can live in one program.</summary>
public class AnnotatedPortal
{
    [XlsxColumn("Name Ex", Order = 3)]
    public string Name { get; set; } = "";
    [XlsxColumn("Owner Ex", Order = 1)]
    public string Owner { get; set; } = "";
    [XlsxColumn("Level Ex", Order = 2)]
    public int Level { get; set; }
    [XlsxIgnore]
    public string InternalNote { get; set; } = "";
}

/// <summary>README Example-7: a ratio shown as a percentage via CustomFormats[0].</summary>
public class PercentSerializer : IXlsxSerializer<double>
{
    public void WriteTitle(XlsxCellWriter writer, double value, XlsxSerializerOptions options, string name = "value")
        => writer.Write(name);
    public void Serialize(XlsxCellWriter writer, double value, XlsxSerializerOptions options)
        => writer.Write(value, customFormat: 0);
}

/// <summary>README Example-6: writes bool cells as YES/NO text.</summary>
public class YesNoSerializer : IXlsxSerializer<bool>
{
    public void WriteTitle(XlsxCellWriter writer, bool value, XlsxSerializerOptions options, string name = "value")
        => writer.Write(name);
    public void Serialize(XlsxCellWriter writer, bool value, XlsxSerializerOptions options)
        => writer.Write(value ? "YES" : "NO");
}

/// <summary>Per-member serializer via [XlsxSerializer], used by the kitchen sink.</summary>
public class UnixSecondsSerializer : IXlsxSerializer<DateTime>
{
    public void WriteTitle(XlsxCellWriter writer, DateTime value, XlsxSerializerOptions options, string name = "value")
        => writer.Write(name);
    public void Serialize(XlsxCellWriter writer, DateTime value, XlsxSerializerOptions options)
        => writer.Write(((DateTimeOffset)value).ToUnixTimeSeconds());
}

#pragma warning disable CS8618
public class Order
{
    public int OrderId { get; set; }
    public string Item { get; set; }
    public int Quantity { get; set; }
    public int? LotNumber { get; set; }
}

public class User
{
    public User(int userId, string ssn)
    {
        Id = userId;
        SSN = ssn;
    }

    public int Id { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string FullName { get; set; }
    public string UserName { get; set; }
    public string Email { get; set; }
    public string SomethingUnique { get; set; }
    public Guid SomeGuid { get; set; }
    public bool SendFlag { get; set; }

    public string Avatar { get; set; }
    public Guid CartId { get; set; }
    public string SSN { get; set; }
    [XlsxSerializer(typeof(UnixSecondsSerializer))]
    public DateTime TimeStamp { get; set; }
    public DateTime CreateTime { get; set; }
    public DateOnly DateOnlyValue { get; set; }
    public TimeOnly TimeOnlyValue { get; set; }
    public TimeSpan TimeSpanValue { get; set; }
    public DateTimeOffset DateTimeOffsetValue { get; set; }
    public object Fallback { get; set; }
    public Uri Uri { get; set; }
    public Bogus.DataSets.Name.Gender Gender { get; set; }

    public List<Order> Orders { get; set; }
    public double DoubleValue { get; set; }
    public char Char { get; set; }
    public string Escape { get; set; }
}
#pragma warning restore CS8618
