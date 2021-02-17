using SimplyExcel.NET;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.Demo
{
    public enum Test
    {
        Demo,
        Demo2,
        Demo3
    }
    public class SampleData
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public SampleDate DateOfBirth { get; set; }
        public string DOB => DateOfBirth.ToString();
        public Test Demo { get; set; }
    }

    public class SampleDate
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public int Day { get; set; }

        public NestedSampleDate Ti { get; set; }
        public override string ToString()
        {
            return $"{Day}/{Month}/{Year}";
        }
    }

    public class NestedSampleDate
    {
        public int Ti { get => 1; }
    }
}
