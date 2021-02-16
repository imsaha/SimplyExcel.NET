using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET.Builder
{
    public class ExcelFileWritingConfiguration
    {
        public int StartingRowIndex { get; set; }
        public CultureInfo Culture { get; set; }

    }
}
