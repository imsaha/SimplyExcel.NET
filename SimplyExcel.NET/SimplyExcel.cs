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
        public IEnumerable<T> Read<T>(Stream stream, Action<ExcelFileReadingConfiguration<T>> fileReadingConfigurationAction = null) where T : class
        {
            return SimplyExcelReader.ReadFromStream(stream, fileReadingConfigurationAction);
        }
    }
}
