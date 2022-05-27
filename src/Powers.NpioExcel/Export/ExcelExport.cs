using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using Powers.NpioExcel;
using Powers.NpioExcel.Attributes;
using Powers.NpioExcel.Interfaces;

#nullable disable
namespace Powers.NpioExcel
{
    public class ExcelExport<T> where T : IExcelStruct
    {
        private HSSFWorkbook _workbook;
        private ISheet _sheet;

        private IRow _headerRow { get; set; }

        private ICellStyle _headerStyle { get; set; }
        private ICellStyle _contentStyle { get; set; }

        private IEnumerable<T> _list { get; set; }

        public ExcelExport(ICellStyle headerStyle = null, ICellStyle contentStyle = null)
        {
            _workbook = new HSSFWorkbook();
            _sheet = _workbook.CreateSheet("Sheet1");

            _headerStyle = headerStyle is null ? ExcelHelper.GetHeaderStyle(_workbook) : headerStyle;
            _contentStyle = contentStyle is null ? ExcelHelper.GetCommonStyle(_workbook) : contentStyle;
        }

        public ExcelExport<T> Export(IEnumerable<T> data)
        {
            _list = data;

            return this;
        }

        private void SetHeader()
        {
            var type = typeof(T);
            var properties = type.GetProperties().OrderBy(x => x.GetCustomAttribute<ExcelColumnAttribute>().Index);

            // var cols_count = properties.Count();

            // var headers_items = properties.Select(x => x.GetCustomAttribute<ExcelColumnAttribute>().Name.Split('/').DistinctBy(y => y).ToArray());

            // var max = headers_items.Select(x => x.Length).Max();

            // var headers = new List<string[]>();

            // // 将headers转为树形结构
            // for (int i = 0; i < max; i++)
            // {
            //     var item = headers_items.Select(x => x.Length > i ? x[i] : "");
            //     headers.Add(item.ToArray());
            // }

            // foreach (var header in headers)
            // {
            //     var row = headers.IndexOf(header);
            //     _headerRow = _sheet.CreateRow(row);

            //     // 分组求出现的次数
            //     var group = header.GroupBy(x => x).Select(x => new ValueTuple<string, int>(x.Key, x.Count()));
            //     var pre = 0;

            //     foreach (var item in group)
            //     {
            //         var col = group.ToList().IndexOf(item);
            //         var cell = _headerRow.CreateCell(pre);
            //         cell.SetCellValue(item.Item1);
            //         cell.CellStyle = _headerStyle;
            //         if (pre != (pre + item.Item2 - 1))
            //             _sheet.AddMergedRegion(new CellRangeAddress(row, row, pre, pre + item.Item2 - 1));
            //         pre += item.Item2;
            //     }
            // }


            foreach (var property in properties)
            {
                var attr = property.GetCustomAttribute<ExcelColumnAttribute>();
                var cell = _headerRow.CreateCell(attr.Index);
                cell.SetCellValue(attr.Name);
                cell.CellStyle = _headerStyle;
                _sheet.SetColumnWidth(attr.Index, attr.Width * 256);
            }

            _headerRow.Height = 30 * 20;
        }

        private void SetBody()
        {
            foreach (var item in _list)
            {
                var row = _sheet.CreateRow(1 + _sheet.LastRowNum);
                var properties = typeof(T).GetProperties().OrderBy(x => x.GetCustomAttribute<ExcelColumnAttribute>().Index);
                foreach (var property in properties)
                {
                    var attr = property.GetCustomAttribute<ExcelColumnAttribute>();
                    var cell = row.CreateCell(attr.Index);
                    var value = property.GetValue(item);
                    if (value != null)
                    {
                        cell.SetCellValue(value.ToString());
                        cell.CellStyle = _contentStyle;
                    }
                }
            }
        }

        public void ToFile(string filePath)
        {
            SetHeader();
            SetBody();

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                _workbook.Write(fs);
            }
        }
    }
}