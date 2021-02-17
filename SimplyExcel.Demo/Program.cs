using SimplyExcel.NET;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace SimplyExcel.Demo
{
    public enum Test
    {
        Demo,
        Demo2,
        Demo3
    }
    class Program
    {
        static void Main(string[] args)
        {
            var data = new List<SampleData>();
            data.Add(new SampleData()
            {
                FirstName = "demo",
                LastName = "a test",
                DateOfBirth = new SampleDate()
                {
                    Year = 1990,
                    Month = 12,
                    Day = 5
                },
                Demo = Test.Demo3
            });

            using (var stream = new FileStream("demo.xlsx", FileMode.OpenOrCreate, FileAccess.Write))
            {
                SimplyExcelWriter.Write(stream, data, options =>
                {
                    options.ColumnMapBuilder = map =>
                    {
                        map.For(c => c.FirstName, 1, "First name")
                            .SetHeaderBackground(Color.Red)
                            .SetCellFormat("0")
                            .SetHeaderTextColor(Color.White);
                        map.For(c => c.LastName, 2);
                        map.For(c => c.DateOfBirth.Year, 3, "Year of");
                        map.For(c => c.DateOfBirth.Month, 4);
                        map.For(c => c.DateOfBirth.Day, 5);
                        map.For(c => c.DOB, 6, "Calc DOB")
                            .SetCellBackground(Color.Green)
                            .SetCellTextColor(Color.White)
                            .SetHeaderBackground(Color.Green)
                            .SetHeaderTextColor(Color.White)
                            .SetCellFormat("m/d/yyyy");
                    };
                });
            }

            Console.WriteLine("Hello World!");
        }
    }
}
