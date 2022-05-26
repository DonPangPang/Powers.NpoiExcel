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

            _testOutputHelper.WriteLine("����\t����\t�Ա�\t����");

            foreach (var item in data)
            {
                _testOutputHelper.WriteLine($"{item.Name}\t{item.Age}\t{item.Gender}\t{item.Born}");
            }
        }

        public class User : IExcelStruct
        {
            [ExcelColumn(Name = "����")]
            public string Name { get; set; }

            [ExcelColumn(Name = "����")]
            public int Age { get; set; }

            [ExcelColumn(Name = "�Ա�")]
            public string Gender { get; set; }

            [ExcelColumn(Name = "����")]
            public DateTime Born { get; set; }
        }
    }
}