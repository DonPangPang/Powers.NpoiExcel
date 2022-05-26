using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using Powers.NpioExcel.Attributes;
using Powers.NpioExcel.Interfaces;

#nullable disable

namespace Powers.NpioExcel.Import
{
    public class ExcelImport
    {
        public string FilePath { get; set; }
        public string SheetName { get; set; }

        private DataTable _dataTable;

        /// <summary>
        /// 初始化文件路径和Sheet名称
        /// </summary>
        /// <param name="filePath">  </param>
        /// <param name="sheetName"> </param>
        public ExcelImport(string filePath, string sheetName = "Sheet1")
        {
            FilePath = filePath;
            SheetName = sheetName;
        }

        /// <summary>
        /// 转为指定类型
        /// </summary>
        /// <typeparam name="T"> </typeparam>
        /// <returns> </returns>
        public List<T> ToList<T>() where T : IExcelStruct
        {
            if (_dataTable is null) ToDataTable();

            var list = new List<T>();
            var type = typeof(T);
            var properties = type.GetProperties();
            foreach (DataRow row in _dataTable.Rows)
            {
                var t = Activator.CreateInstance<T>();
                foreach (var property in properties)
                {
                    var prop_name = property.GetCustomAttribute<ExcelColumnAttribute>().Name;
                    var value = row[prop_name];
                    if (value != DBNull.Value)
                    {
                        var d = Convert.ChangeType(value, property.PropertyType);
                        property.SetValue(t, d);
                    }
                }
                list.Add(t);
            }
            return list;
        }

        /// <summary>
        /// Excel转为DataTable
        /// </summary>
        /// <returns> </returns>
        private ExcelImport ToDataTable()
        {
            _dataTable = new DataTable();
            using (var fs = new FileStream(FilePath, FileMode.Open, FileAccess.Read))
            {
                var workbook = WorkbookFactory.Create(fs);
                var sheet = workbook.GetSheet(SheetName);
                if (sheet == null)
                {
                    throw new Exception("Excel sheet not found");
                }

                var headerRow = sheet.GetRow(0);
                var headerRowCount = headerRow.LastCellNum;
                for (var i = headerRow.FirstCellNum; i < headerRowCount; i++)
                {
                    var column = new DataColumn(headerRow.GetCell(i).StringCellValue);
                    _dataTable.Columns.Add(column);
                }

                var rowCount = sheet.LastRowNum;
                for (var i = sheet.FirstRowNum + 1; i <= sheet.LastRowNum; i++)
                {
                    var row = sheet.GetRow(i);
                    var dataRow = _dataTable.NewRow();
                    for (var j = row.FirstCellNum; j < headerRowCount; j++)
                    {
                        if (row.GetCell(j) != null)
                        {
                            dataRow[j] = row.GetCell(j).ToString();
                        }
                    }

                    _dataTable.Rows.Add(dataRow);
                }
            }

            return this;
        }
    }
}