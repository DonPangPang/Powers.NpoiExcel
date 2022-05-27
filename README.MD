# Powers.NpoiExcel

Now Support: Import/Export Excel to Entity

1: Create your entity

```csharp
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
```

2: How to use

```csharp
var root = Environment.CurrentDirectory + "/files";

var path = root + "/test.xlsx";

// Import
var data = new ExcelImport(path).ToList<User>();

_testOutputHelper.WriteLine("姓名\t年龄\t性别\t生日");

foreach (var item in data)
{
    _testOutputHelper.WriteLine($"{item.Name}\t{item.Age}\t{item.Gender}\t{item.Born}");
}

var out_path = root + "/out.xlsx";
// Export
new ExcelExport<User>().Export(data).ToFile(out_path);
```

3: Result or run xUnit.Test

```terminal
 Powers.NpioExcel.TestProject.ExcelTests.Test1
   持续时间: 244 毫秒

  标准输出: 
    姓名	年龄	性别	生日
    张三	20	男	2022/5/1 0:00:00
    李四	21	女	2022/5/2 0:00:00
    王五	22	男	2022/5/3 0:00:00
    赵六	23	女	2022/5/4 0:00:00
    田七	24	男	2022/5/5 0:00:00
    老八	25	女	2022/5/6 0:00:00
    小汉堡	26	男	2022/5/7 0:00:00
    奥里给	27	女	2022/5/8 0:00:00
    嘿嘿嘿	28	男	2022/5/9 0:00:00
```
