using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET.Builder
{
    public class ExcelColumnConfiguration
    {
        public string ColumnName { get; set; }
        public int ColumnIndex { get; set; }
        public PropertyInfo Property { get; set; }

    }
}
