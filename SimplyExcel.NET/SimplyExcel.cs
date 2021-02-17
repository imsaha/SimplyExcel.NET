using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SimplyExcel.NET.Builder;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET
{
    public class SimplyExcel : ISimplyExcel
    {
        public IEnumerable<T> Read<T>(Stream stream, Action<ExcelReadingConfiguration<T>> excelConfigurationAction = null) where T : class
        {
            return SimplyExcelReader.ReadFromStream(stream, excelConfigurationAction);
        }

        public void Write<T>(Stream stream, IEnumerable<T> items, Action<ExcelWritingConfiguration<T>> excelConfigurationAction = null) where T : class
        {
            SimplyExcelWriter.Write(stream, items, excelConfigurationAction);
        }
    }
}
