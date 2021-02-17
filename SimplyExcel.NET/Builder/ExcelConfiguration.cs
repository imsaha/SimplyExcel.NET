using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET.Builder
{
    public abstract class ExcelConfiguration<T> where T : class
    {
        public string SheetName { get; set; }
        public int StartingRowIndex { get; set; }
        public int? HeaderRowIndex { get; set; }
        public CultureInfo Culture { get; set; }

        public Action<ExcelColumMapBuilder<T, ExcelColumnConfiguration>> ColumnMapBuilder { get; set; }

        public abstract IReadOnlyCollection<ExcelColumnConfiguration> GetMaps();
    }

    public class ExcelReadingConfiguration<T> : ExcelConfiguration<T> where T : class
    {
        new public Action<ExcelColumMapBuilder<T, ExcelOnReadingColumnConfiguration>> ColumnMapBuilder { get; set; }

        public override IReadOnlyCollection<ExcelColumnConfiguration> GetMaps()
        {
            var columnMapsObj = new ExcelColumMapBuilder<T, ExcelOnReadingColumnConfiguration>();
            ColumnMapBuilder?.Invoke(columnMapsObj);
            return columnMapsObj.GetMaps();
        }
    }

    public class ExcelWritingConfiguration<T> : ExcelConfiguration<T> where T : class
    {
        new public Action<ExcelColumMapBuilder<T, ExcelOnWritingColumnConfiguration>> ColumnMapBuilder { get; set; }

        public override IReadOnlyCollection<ExcelColumnConfiguration> GetMaps()
        {
            var columnMapsObj = new ExcelColumMapBuilder<T, ExcelOnWritingColumnConfiguration>();
            ColumnMapBuilder?.Invoke(columnMapsObj);
            return columnMapsObj.GetMaps();
        }
    }
}
