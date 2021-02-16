using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET.Builder
{
    public class ExcelFileReadingConfiguration<T>
    {
        public int StartingRowIndex { get; set; }
        public CultureInfo Culture { get; set; }
        public Action<ExcelColumnIndexMap<T>> ColumnIndexMap { get; set; }

        public Dictionary<PropertyInfo, int> GetColumnMaps()
        {
            var columnMapsObj = new ExcelColumnIndexMap<T>();
            if (ColumnIndexMap != null)
                ColumnIndexMap(columnMapsObj);

            return columnMapsObj.GetIndexMap();
        }
    }
}
