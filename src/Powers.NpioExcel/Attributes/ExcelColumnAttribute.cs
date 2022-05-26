using System;
using System.Collections.Generic;
using System.Text;

namespace Powers.NpioExcel.Attributes
{
    public class ExcelColumnAttribute : Attribute
    {
        public string Name { get; set; } = null!;

        public int Index { get; set; } = 0;
    }
}