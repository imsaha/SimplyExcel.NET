using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET.Builder
{
    public abstract class ExcelColumnConfiguration
    {
        public string ColumnName { get; set; }
        public int ColumnIndex { get; set; }
        public PropertyInfo Property { get; set; }
        public ExcelColumnConfiguration[] Children { get; set; }
    }

    public class ExcelOnReadingColumnConfiguration : ExcelColumnConfiguration
    {

    }

    public class ExcelOnWritingColumnConfiguration : ExcelColumnConfiguration
    {
        private readonly ExcelHeaderConfiguration _header;
        private readonly ExcelCellConfiguration _cell;

        public ExcelOnWritingColumnConfiguration()
        {
            _header = new ExcelHeaderConfiguration();
            _cell = new ExcelCellConfiguration();
        }

        public ExcelHeaderConfiguration Header => _header;
        public ExcelCellConfiguration Cell => _cell;


        public ExcelOnWritingColumnConfiguration SetHeaderBackground(Color color)
        {
            _header.SetBackground(color);
            return this;
        }

        public ExcelOnWritingColumnConfiguration SetHeaderTextColor(Color color)
        {
            _header.SetTextColor(color);
            return this;
        }

        public ExcelOnWritingColumnConfiguration SetHeaderFormat(string format)
        {
            _header.SetFormat(format);
            return this;
        }


        public ExcelOnWritingColumnConfiguration SetCellBackground(Color color)
        {
            _cell.SetBackground(color);
            return this;
        }

        public ExcelOnWritingColumnConfiguration SetCellTextColor(Color color)
        {
            _cell.SetTextColor(color);
            return this;
        }

        public ExcelOnWritingColumnConfiguration SetCellFormat(string format)
        {
            _cell.SetFormat(format);
            return this;
        }
    }


    public class ExcelCellConfiguration
    {
        private Color _textColor;
        private Color _backgroundColor;
        private string _format;

        public Color BackgroundColor { get => _backgroundColor; }
        public Color TextColor { get => _textColor; }
        public string Format { get => _format; }

        public ExcelCellConfiguration SetBackground(Color color)
        {
            _backgroundColor = color;
            return this;
        }

        public ExcelCellConfiguration SetTextColor(Color color)
        {
            _textColor = color;
            return this;
        }

        public ExcelCellConfiguration SetFormat(string format)
        {
            _format = format;
            return this;
        }
    }

    public class ExcelHeaderConfiguration : ExcelCellConfiguration
    {

    }
}
