using Powers.NpioExcel.Attributes;
using Powers.NpioExcel.Import;
using Powers.NpioExcel.Interfaces;
using Xunit.Abstractions;

namespace Powers.NpioExcel.TestProject
{
    public class ExcelTests
    {
        private readonly ITestOutputHelper _testOutputHelper;

        public ExcelTests(ITestOutputHelper testOutputHelper)
        {
            _testOutputHelper = testOutputHelper;
        }

        [Fact]
        public void Test1()
        {
            var path = Environment.CurrentDirectory + "/files/test.xlsx";

            var data = new ExcelImport(path).ToList<User>();

            _testOutputHelper.WriteLine("姓名\t年龄\t性别\t生日");

            foreach (var item in data)
            {
                _testOutputHelper.WriteLine($"{item.Name}\t{item.Age}\t{item.Gender}\t{item.Born}");
            }
        }

        public class User : IExcelStruct
        {
            [ExcelColumn(Name = "姓名")]
            public string Name { get; set; }

            [ExcelColumn(Name = "年龄")]
            public int Age { get; set; }

            [ExcelColumn(Name = "性别")]
            public string Gender { get; set; }

            [ExcelColumn(Name = "生日")]
            public DateTime Born { get; set; }
        }
    }
}