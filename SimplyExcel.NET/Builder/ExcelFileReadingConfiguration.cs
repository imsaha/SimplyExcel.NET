using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET.Builder
{
    public class ExcelConfiguration<T> where T : class
    {
        public int StartingRowIndex { get; set; }
        public int? HeaderRowIndex { get; set; }
        public CultureInfo Culture { get; set; }
        public Action<ExcelColumMapBuilder<T>> MapBuilderAction { get; set; }

        public IReadOnlyCollection<ExcelColumnConfiguration> GetColumnMaps()
        {
            var columnMapsObj = new ExcelColumMapBuilder<T>();
            MapBuilderAction?.Invoke(columnMapsObj);

            return columnMapsObj.GetIndexMap();
        }
    }

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
