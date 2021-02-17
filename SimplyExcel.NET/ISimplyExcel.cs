using SimplyExcel.NET.Builder;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET
{
    public interface ISimplyExcel
    {
        IEnumerable<T> Read<T>(Stream stream, Action<ExcelReadingConfiguration<T>> excelConfigurationAction = null) where T : class;
        void Write<T>(Stream stream, IEnumerable<T> items, Action<ExcelWritingConfiguration<T>> excelConfigurationAction = null) where T : class;
    }
}
