using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using SimplyExcel.NET.Builder;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SimplyExcel.NET
{
    public static class SimplyExcelWriter
    {
        public static void Write<T>(Stream stream, IEnumerable<T> items, Action<ExcelWritingConfiguration<T>> excelMapBuilderAction = null) where T : class
        {
            if (items == null)
                throw new NullReferenceException("expected parameter is null.");

            var configuration = new ExcelWritingConfiguration<T>();

            if (excelMapBuilderAction != null)
                excelMapBuilderAction(configuration);

            var xlBook = new XSSFWorkbook();
            var sheet1 = xlBook.CreateSheet(configuration.SheetName ?? "Sheet1");

            var maps = (IReadOnlyCollection<ExcelOnWritingColumnConfiguration>)configuration.GetMaps();
            createHeader(items.FirstOrDefault(), xlBook, sheet1, 0, maps);
            var headerEndIndex = sheet1.LastRowNum;

            for (int i = 0; i < items.Count(); i++)
            {
                var item = items.ElementAt(i);
                createRow(xlBook, sheet1, item, maps);
            }

            buildDataValidationForEnum<T>(sheet1, headerEndIndex, maps);

            xlBook.Write(stream);
        }

        private static XSSFCellStyle createCellStyle(XSSFWorkbook xlBook, bool isHeader = false)
        {
            var cellStyle = (XSSFCellStyle)xlBook.CreateCellStyle();
            cellStyle.Alignment = isHeader ? HorizontalAlignment.Center : HorizontalAlignment.Left;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;

            var headerFont = xlBook.CreateFont();
            headerFont.IsBold = isHeader;
            cellStyle.SetFont(headerFont);

            return cellStyle;
        }

        private static void createHeader<T>(T item, XSSFWorkbook xlBook, ISheet sheet, int columnIndex, IEnumerable<ExcelOnWritingColumnConfiguration> columnMaps, CellRangeAddress mergedCellRange = null) where T : class
        {
            var headerRow = sheet.CreateRow(sheet.LastRowNum + 1);
            foreach (var column in columnMaps)
            {
                var headerConfig = column.Header;
                var headerStyle = createCellStyle(xlBook, true);

                var cell = headerRow.CreateCell(columnIndex + column.ColumnIndex, CellType.String);
                cell.CellStyle = headerStyle;
                cell.SetCellValue(column.ColumnName);

                if (!headerConfig.BackgroundColor.IsEmpty)
                {
                    var backgroundColor = new XSSFColor(headerConfig.BackgroundColor);
                    headerStyle.SetFillForegroundColor(backgroundColor);
                    headerStyle.FillPattern = FillPattern.SolidForeground;
                }

                if (!headerConfig.TextColor.IsEmpty)
                {
                    var textColor = new XSSFColor(headerConfig.TextColor);
                    headerStyle.GetFont().SetColor(textColor);
                }

                // add merged cell
                CellRangeAddress mergedArea = null;
                if (column.Children != null)
                {
                    // add merged cell horizontally
                    mergedArea = new CellRangeAddress(headerRow.RowNum, headerRow.RowNum, column.ColumnIndex, (column.ColumnIndex + column.Children.Length - 1));
                    createHeader(item, xlBook, sheet, column.ColumnIndex, ((ExcelOnWritingColumnConfiguration[])column.Children).ToList(), mergedArea);
                }
                else
                {
                    // add merged cell vertically
                    if (columnMaps.Any(x => x.Children != null && x.Children.Count() > 0))
                        mergedArea = new CellRangeAddress(headerRow.RowNum, headerRow.RowNum + 1, column.ColumnIndex, column.ColumnIndex);
                }

                if (mergedArea != null)
                    sheet.AddMergedRegion(mergedArea);
            }
        }

        private static void createRow<T>(XSSFWorkbook xlBook, ISheet sheet, T item, IEnumerable<ExcelOnWritingColumnConfiguration> columnMaps)
        {
            var lastRow = sheet.LastRowNum + 1;
            var row = sheet.CreateRow(lastRow);
            var columnIndex = 0;
            for (int i = 0; i < columnMaps.Count(); i++)
            {
                var column = columnMaps.ElementAt(i);
                var isSubProp = column.Children != null && column.Children.Count() > 0;
                if (isSubProp)
                {
                    var subItem = column.Property.GetValue(item);
                    foreach (var child in column.Children)
                    {
                        createCell(child.Property.GetValue(subItem), (child as ExcelOnWritingColumnConfiguration), xlBook, row, columnIndex++);
                    }
                }
                else
                {
                    createCell(column.Property.GetValue(item), column, xlBook, row, columnIndex++);
                }
            }
        }

        private static void createCell<T>(T value, ExcelOnWritingColumnConfiguration configuration, XSSFWorkbook xlBook, IRow row, int columnIndex = 0)
        {
            var cell = row.CreateCell(columnIndex);
            var cellStyle = (XSSFCellStyle)xlBook.CreateCellStyle();
            cell.CellStyle = cellStyle;

            var property = configuration.Property;

            if (property.PropertyType == typeof(int) ||
                property.PropertyType == typeof(int?) ||
                property.PropertyType == typeof(long) ||
                property.PropertyType == typeof(long?) ||
                property.PropertyType == typeof(double) ||
                property.PropertyType == typeof(double?) ||
                property.PropertyType == typeof(float) ||
                property.PropertyType == typeof(float?) ||
                property.PropertyType == typeof(decimal) ||
                property.PropertyType == typeof(decimal?))
            {
                var dataFormat = xlBook.CreateDataFormat();
                cellStyle.DataFormat = dataFormat.GetFormat("_(* #,##0_);_(* (#,##0);_(* \" - \"_);_(@_)");
                if (value != null)
                    cell.SetCellValue(Convert.ToDouble(value));
            }
            else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
            {
                var dataFormat = xlBook.CreateDataFormat();
                cellStyle.DataFormat = dataFormat.GetFormat("d-mmm-yy");

                if (value != null)
                    cell.SetCellValue(DateTime.Parse(value.ToString()));
            }
            else if (property.PropertyType == typeof(Enum))
            {
                if (value != null)
                    cell.SetCellValue(value.ToString());
            }
            else
            {
                if (value != null)
                    cell.SetCellValue(value.ToString());
            }


            var font = xlBook.CreateFont();
            font.IsBold = false;

            var cellConfig = configuration.Cell;
            if (!cellConfig.BackgroundColor.IsEmpty)
            {
                var backgroundColor = new XSSFColor(cellConfig.BackgroundColor);
                cellStyle.SetFillForegroundColor(backgroundColor);
                cellStyle.FillPattern = FillPattern.SolidForeground;
            }

            if (!cellConfig.TextColor.IsEmpty)
            {
                var textColor = new XSSFColor(cellConfig.TextColor);
                cellStyle.GetFont().SetColor(textColor);
            }

            if (!string.IsNullOrWhiteSpace(cellConfig.Format))
            {
                var dataFormat = xlBook.CreateDataFormat();
                cellStyle.DataFormat = dataFormat.GetFormat(cellConfig.Format);
            }

            cellStyle.SetFont(font);
        }

        private static void buildDataValidationForEnum<T>(ISheet sheet1, int headerEndIndex, IEnumerable<ExcelOnWritingColumnConfiguration> columnMaps)
        {
            var cellIndex = 0;
            for (int i = 0; i < columnMaps.Count(); i++)
            {
                var columnMap = columnMaps.ElementAt(i);
                var property = columnMap.Property;
                var isRefType = property.PropertyType.IsClass && property.PropertyType.Assembly != typeof(object).Assembly;
                if (isRefType)
                    cellIndex = (cellIndex - 1) + (property.PropertyType.GetProperties().Length);

                if (property.PropertyType.IsEnum)
                {
                    var names = EnumHelper.GetNames(property.PropertyType);
                    if (names == null || names.Count == 0)
                        continue;

                    var dvHelper = new XSSFDataValidationHelper((XSSFSheet)sheet1);
                    var constraint = (XSSFDataValidationConstraint)dvHelper.CreateExplicitListConstraint(names.ToArray());
                    var addressList = new CellRangeAddressList(headerEndIndex, sheet1.LastRowNum, cellIndex, cellIndex);
                    var validation = (XSSFDataValidation)dvHelper.CreateValidation(constraint, addressList);

                    validation.EmptyCellAllowed = true;
                    validation.SuppressDropDownArrow = true;
                    validation.ShowErrorBox = true;
                    validation.ErrorStyle = 0;//Stop
                    validation.CreateErrorBox("Invalid selection", "Select a valid data from the list.");

                    sheet1.AddValidationData(validation);
                }
                cellIndex++;
            }
        }

    }
}
