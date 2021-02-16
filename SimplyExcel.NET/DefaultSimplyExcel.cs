﻿using NPOI.SS.UserModel;
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
    public class DefaultSimplyExcel : ISimplyExcel
    {
        public IEnumerable<T> Read<T>(Stream stream, Action<ExcelFileReadingConfiguration<T>> fileReadingConfigurationAction = null) where T : class
        {
            if (stream == null || stream.Length == 0)
                throw new ArgumentNullException(nameof(stream));

            var configuration = new ExcelFileReadingConfiguration<T>();
            fileReadingConfigurationAction(configuration);
            var columnMaps = configuration.GetColumnMaps();

            var result = new List<T>();
            stream.Position = 0;
            var book = new XSSFWorkbook(stream);
            for (int sheetIndex = 0; sheetIndex < book.NumberOfSheets; sheetIndex++)
            {
                var sheet = book.GetSheetAt(sheetIndex);
                for (int rowIndex = (sheet.FirstRowNum + configuration.StartingRowIndex); rowIndex < (sheet.LastRowNum + 1); rowIndex++)
                {
                    var instanceObj = Activator.CreateInstance<T>();
                    var row = sheet.GetRow(rowIndex);
                    for (int matchIndex = 0; matchIndex < columnMaps.Count; matchIndex++)
                    {
                        var match = columnMaps.ElementAt(matchIndex);
                        var property = match.Key;
                        var cell = row.GetCell(match.Value, MissingCellPolicy.RETURN_NULL_AND_BLANK);
                        var cellData = cell?.ToString();
                        if (cell == null || string.IsNullOrEmpty(cellData))
                        {
                            property.SetValue(instanceObj, default);
                            continue;
                        }

                        if (property.PropertyType == typeof(DateTime))
                        {
                            var isParsed = DateTime.TryParse(cellData, configuration.Culture ?? CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out DateTime parsedDateTime);
                            property.SetValue(instanceObj, parsedDateTime);
                            if (!isParsed)
                                throw new ArgumentException("Invalid date value = " + cellData);

                            continue;
                        }

                        TypeConverter converter = TypeDescriptor.GetConverter(property.PropertyType);
                        if (converter.CanConvertTo(property.PropertyType) && converter.CanConvertFrom(typeof(string)))
                        {
                            var convertedData = converter.ConvertFromString(cellData);
                            property.SetValue(instanceObj, convertedData);
                            continue;
                        }

                        throw new ArgumentException("Data = " + cellData + " can not convert.");
                    }

                    result.Add(instanceObj);
                }
            }

            return result;
        }
    }
}