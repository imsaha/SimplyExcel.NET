# SimplyExcel.NET


```c#
  List<SampleData> data = new List<SampleData>
            {
                new SampleData()
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
                },
                new SampleData()
                {
                    FirstName = "demo2",
                    LastName = "an another test",
                    DateOfBirth = new SampleDate()
                    {
                        Year = 1986,
                        Month = 06,
                        Day = 9
                    },
                    Demo = Test.Demo3
                }
            };

```


### export simple excel 


```c#

        static void buildSimple(List<SampleData> data, string fileName)
        {
            using (var stream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write))
            {
                SimplyExcelWriter.Write(stream, data);
            }
        }


```

### export excel with advance configuration 

```c#

        static void buildAdvance(List<SampleData> data, string fileName)
        {
            using (var stream = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.Write))
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

        }


```


