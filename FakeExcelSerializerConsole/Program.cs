using Bogus;
using FakeExcelSerializer;
using System.Diagnostics;
using System.Globalization;
using static Bogus.DataSets.Name;

var sw = Stopwatch.StartNew();

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

var Users = testUsers.Generate(10);

sw.Stop();
Console.WriteLine($"testUsers.Generate count:{Users.Count:#,##0} duration:{sw.ElapsedMilliseconds:#,##0}ms");
sw.Restart();

var newConfig = ExcelSerializerOptions.Default with
{
    CultureInfo = CultureInfo.CurrentCulture,
    MaxDepth = 32,
    Provider = ExcelSerializerProvider.Create(
        //new[] { new BoolZeroOneSerializer() },
        new[] { ExcelSerializerProvider.Default }),
    HasHeaderRecord = true,
    AutoFitColumns = true,
    AutoFilter = true,
};

var fileName = Path.Combine(Environment.CurrentDirectory, "test.xlsx");
if (File.Exists(fileName))
    File.Delete(fileName);

ExcelSerializer.ToFile(Users, fileName, newConfig);

sw.Stop();

Console.WriteLine($"ExcelSerializer.ToFile duration:{sw.ElapsedMilliseconds:#,##0}ms");
Console.WriteLine($"Excel file created. Please check the file. {fileName}");

Console.WriteLine();
Console.WriteLine("press any key...");

Console.ReadLine();


//public class BoolZeroOneSerializer : IExcelSerializer<bool>
//{
//    public void WriteTitle(ref ExcelSerializerWriter writer, bool value, ExcelSerializerOptions options, string name = "")
//    {
//        writer.Write(name);
//    }

//    public void Serialize(ref ExcelSerializerWriter writer, bool value, ExcelSerializerOptions options)
//    {
//        // true => 0, false => 1
//        writer.WritePrimitive(value ? 0 : 1);
//    }
//}

public class UnixSecondsSerializer : IExcelSerializer<DateTime>
{
    public void WriteTitle(ref ExcelSerializerWriter writer, DateTime value, ExcelSerializerOptions options, string name = "")
    {
        writer.Write(name);
    }

    public void Serialize(ref ExcelSerializerWriter writer, DateTime value, ExcelSerializerOptions options)
    {
        writer.WritePrimitive(((DateTimeOffset)(value)).ToUnixTimeSeconds());
    }
}

#pragma warning disable CS8618 // null 非許容のフィールドには、コンストラクターの終了時に null 以外の値が入っていなければなりません。Null 許容として宣言することをご検討ください。
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
    [ExcelSerializer(typeof(UnixSecondsSerializer))]
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
#pragma warning restore CS8618 // null 非許容のフィールドには、コンストラクターの終了時に null 以外の値が入っていなければなりません。Null 許容として宣言することをご検討ください。
