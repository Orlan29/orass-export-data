using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using Aspose.Cells;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using InfiSoftware.Common.DataAccess.SpreadsheetGeneration.Exceptions;
using InfiSoftware.Common.DataAccess.SpreadsheetGeneration.Model;
using InfiSoftware.Core.Utilities.Utility;

namespace InfiSoftware.Common.DataAccess.SpreadsheetGeneration.Excel
{
    public class ExcelGeneration : ISpreadsheetGeneration
    {
        private class QuickFillInfo
        {
            public int CurrentRow { get; set; }
            public int CurrentColumn { get; set; }
            public int LastColumnExtend { get; set; }
        }

        private ExcelPackage _package;
        private ExcelWorksheet _activeSheet;
        private readonly Color _red = Color.FromArgb(255, 0, 0);
        private readonly Color _yellow = Color.FromArgb(255, 255, 0);
        private readonly Dictionary<string, QuickFillInfo> _quickFill = new Dictionary<string, QuickFillInfo>();
        public static bool UseAsPoseCalc { get; set; } = true;
        private const string DoubleConversionFormat = "0.###################################################################################################################################################################################################################################################################################################################################################";

        static ExcelGeneration()
        {
            using (var stream = EmbeddedResource.ReadResourceFromAsStream(typeof(ExcelGeneration), "SpreadsheetGeneration/Excel/Licenses/Aspose.Cells.lic"))
            {
                var lic = new License();
                lic.SetLicense(stream);
            }
        }

        public ExcelGeneration()
        {
            _package = new ExcelPackage();
        }

        public ExcelGeneration(string filePath, string password = null)
        {
            var excelFileInfo = new FileInfo(filePath);
            if (!excelFileInfo.Exists)
                throw new SpreadsheetGenerationExceptions.OpenFileException("Cannot open file");
            _package = new ExcelPackage(excelFileInfo, password);
            _activeSheet = _package.Workbook.Worksheets.FirstOrDefault(f => f.View.TabSelected);
        }

        public ExcelGeneration(byte[] fileBytes) : this(new MemoryStream(fileBytes))
        {
        }

        public ExcelGeneration(Stream content)
        {
            _package = new ExcelPackage(content);
            _activeSheet = _package.Workbook.Worksheets.FirstOrDefault();
        }

        public ExcelGeneration(string worksheetName, params object[] columns) : this()
        {
            _package = new ExcelPackage();
            AddQuickWorksheet(worksheetName, columns);
        }

        public ExcelGeneration(string worksheetName, params FusionCell[] columns) : this()
        {
            _package = new ExcelPackage();
            AddQuickWorksheetWithMerge(worksheetName, columns);
        }

        public void AddQuickWorksheet(string worksheetName, params object[] columns)
        {
            AddAndSelectWorkSheet(worksheetName);
            _quickFill[worksheetName] = new QuickFillInfo { CurrentRow = 1, LastColumnExtend = columns.Length + 1 };
            for (var colum = 0; colum < columns.Length; colum++)
                SetCellValue(columns[colum], colum);
        }

        public void AddQuickWorksheetWithMerge(string worksheetName, params FusionCell[] columns)
        {
            AddAndSelectWorkSheet(worksheetName);
            _quickFill[worksheetName] = new QuickFillInfo { CurrentRow = 1, LastColumnExtend = columns.Sum(c => c.CellWidth) + 1 };
            for (var column = 0; column < columns.Length; column++)
                SetCellValue(columns[column].CellValue, column > 0 ? columns.Take(column).Sum(c => c.CellWidth) : 0, columns[column].CellWidth);
        }

        public void QuickFillRow(params object[] rowCells)
        {
            if (!_quickFill.ContainsKey(_activeSheet.Name))
                _quickFill[_activeSheet.Name] = new QuickFillInfo { CurrentRow = 1, LastColumnExtend = rowCells.Length + 1 };
            _quickFill[_activeSheet.Name].CurrentRow++;
            for (var column = 0; column < rowCells.Length; column++)
                SetCellValue(rowCells[column], column);
            _quickFill[_activeSheet.Name].CurrentColumn = rowCells.Length;
        }

        public void QuickFillRowWithMerge(params FusionCell[] rowCells)
        {
            if (!_quickFill.ContainsKey(_activeSheet.Name))
                _quickFill[_activeSheet.Name] = new QuickFillInfo { CurrentRow = 1, LastColumnExtend = rowCells.Sum(c => c.CellWidth) + 1 };
            _quickFill[_activeSheet.Name].CurrentRow++;
            for (var column = 0; column < rowCells.Length; column++)
                SetCellValue(rowCells[column].CellValue, column > 0 ? rowCells.Take(column).Sum(c => c.CellWidth) : 0, rowCells[column].CellWidth);
            _quickFill[_activeSheet.Name].CurrentColumn = rowCells.Sum(c => c.CellWidth);
        }

        public void QuickFillRowByList(IList<string> rowCells)
        {
            if (!_quickFill.ContainsKey(_activeSheet.Name))
                _quickFill[_activeSheet.Name] = new QuickFillInfo { CurrentRow = 1, LastColumnExtend = rowCells.Count + 1 };
            _quickFill[_activeSheet.Name].CurrentRow++;
            for (var column = 0; column < rowCells.Count; column++)
                SetCellValue(rowCells[column], column);
            _quickFill[_activeSheet.Name].CurrentColumn = rowCells.Count;
        }

        public void QuickFillRowAppend(params object[] rowCells)
        {
            if (!_quickFill.ContainsKey(_activeSheet.Name))
                _quickFill[_activeSheet.Name] = new QuickFillInfo { CurrentRow = 2, LastColumnExtend = rowCells.Length + 1 };
            foreach (var rowCell in rowCells)
                SetCellValue(rowCell, _quickFill[_activeSheet.Name].CurrentColumn++);
            _quickFill[_activeSheet.Name].LastColumnExtend = Math.Max(_quickFill[_activeSheet.Name].CurrentColumn, _quickFill[_activeSheet.Name].LastColumnExtend);
        }

        public void SelectWorksheet(object worksheet)
        {
            _activeSheet = worksheet is string index ? _package.Workbook.Worksheets[index] : _package.Workbook.Worksheets[(int)worksheet];
        }

        public void DeleteWorksheet(string worksheetName)
        {
            var worksheet = _package.Workbook.Worksheets.SingleOrDefault(x => x.Name == worksheetName);
            _package.Workbook.Worksheets.Delete(worksheet);
        }

        public void AddWorkSheet(string worksheetName)
        {
            _package.Workbook.Worksheets.Add(worksheetName);
        }

        public void AddAndSelectWorkSheet(string worksheetName)
        {
            var ws = _package.Workbook.Worksheets.FirstOrDefault(x => x.Name == worksheetName);
            if (ws == null)
                _package.Workbook.Worksheets.Add(worksheetName);
            _activeSheet = _package.Workbook.Worksheets[worksheetName];
        }

        public bool WorksheetExist()
        {
            return _activeSheet != null;
        }

        public bool WorksheetExist(string worksheetName)
        {
            return _package.Workbook.Worksheets.Any(m => m.Name == worksheetName);
        }

        public void Save(string password = null)
        {
            foreach (var worksheet in _package.Workbook.Worksheets)
                if (_quickFill.ContainsKey(worksheet.Name))
                    QuickFilterAndFreeze(worksheet.Name);
            _package.Save(password);
        }

        public void SaveToFile(string filePath, string password = null)
        {
            foreach (var worksheet in _package.Workbook.Worksheets)
                if (_quickFill.ContainsKey(worksheet.Name))
                    QuickFilterAndFreeze(worksheet.Name);
            _package.SaveAs(new FileInfo(filePath), password);
        }

        public void SaveToFileStamp(string filePath, string password = null)
        {
            var now = DateTime.Now;
            var stamp = $" {now.Year}-{now.Month:00}-{now.Day:00} à {now.Hour:00}h{now.Minute:00}";
            SaveToFile(filePath + stamp + ".xlsx", password);
        }

        public void SaveToCsvFile(string filePath, object worksheet, string delimiter = ",", bool hasQuotesTextQualifier = false)
        {
            var bytes = SaveToCsvBytes(worksheet, delimiter, hasQuotesTextQualifier);
            File.WriteAllBytes(filePath, bytes);
        }

        public byte[] SaveToBytes(string password = null, bool freezePanels = true)
        {
            foreach (var worksheet in _package.Workbook.Worksheets)
                if (_quickFill.ContainsKey(worksheet.Name))
                    QuickFilterAndFreeze(worksheet.Name, freezePanels);
            return _package.GetAsByteArray(password);
        }

        public void SaveToCsvStream(Stream outStream, object worksheet, string delimiter = ",", bool hasQuotesTextQualifier = false)
        {
            var sheet = worksheet is string index ? _package.Workbook.Worksheets[index] : _package.Workbook.Worksheets[(int)worksheet];
            var maxColumnNumber = sheet.Dimension.End.Column;
            var currentRow = new List<string>(maxColumnNumber);
            var totalRowCount = sheet.Dimension.End.Row;
            var currentRowNum = 1;
            using (var writer = new StreamWriter(outStream, Encoding.UTF8, 1024, true))
            {
                while (currentRowNum <= totalRowCount)
                {
                    for (var i = 1; i <= maxColumnNumber; i++)
                    {
                        var cell = sheet.Cells[currentRowNum, i];
                        currentRow.Add(string.Format("{0}{1}{0}", hasQuotesTextQualifier ? "\"" : string.Empty, cell == null ? string.Empty : cell.Value?.ToString() ?? string.Empty));
                    }
                    var delimitedString = ToDelimitedString(currentRow, delimiter);
                    if (currentRowNum == totalRowCount)
                        writer.Write(delimitedString);
                    else
                        writer.WriteLine(delimitedString);
                    currentRow.Clear();
                    currentRowNum++;
                }
            }
        }

        public byte[] SaveToCsvBytes(object worksheet, string delimiter = ",", bool hasQuotesTextQualifier = false)
        {
            using (var memory = new MemoryStream())
            {
                this.SaveToCsvStream(memory, worksheet, delimiter, hasQuotesTextQualifier);
                return memory.ToArray();
            }
        }

        public void ImportCsvFile(string filePath, string delimiter = ",", bool hasQuotesTextQualifier = false)
        {
            var format = new ExcelTextFormat
            {
                Delimiter = delimiter[0],
                Encoding = new UTF8Encoding(),
                Culture = new CultureInfo(System.Threading.Thread.CurrentThread.CurrentCulture.ToString()) { DateTimeFormat = { ShortDatePattern = "dd.mm.yyyy" } }
            };
            if (hasQuotesTextQualifier)
                format.TextQualifier = '"';
            _activeSheet.Cells["A1"].LoadFromText(new FileInfo(filePath), format);
        }

        public void FinishFormatting(int lastColumn)
        {
            for (var index = 1; index < lastColumn; index++)
                SetColumnAutoWidth(index);
            SetWorksheetAutofilter();
            FreezeWorksheetPanes(2, 1);
        }

        public void FinishFormatting(int lastColumn, string startFilterCell)
        {
            for (var index = 1; index < lastColumn; index++)
                SetColumnAutoWidth(index);
            SetWorksheetAutofilter(startFilterCell, null);
            FreezeWorksheetPanes(2, 1);
        }

        public void SetWorksheetAutofilter()
        {
            _activeSheet.Cells[_activeSheet.Dimension.Address].AutoFilter = true;
        }

        public void SetWorksheetAutofilter(string from, string to)
        {
            if (string.IsNullOrEmpty(from))
                from = _activeSheet.Dimension.Start.Address;
            if (string.IsNullOrEmpty(to))
                to = _activeSheet.Dimension.End.Address;

            _activeSheet.Cells[from + ":" + to].AutoFilter = true;
        }

        public void FreezeWorksheetPanes(int row, object column)
        {
            _activeSheet.View.FreezePanes(row, column.ToColumnInt());
        }

        public void SetRowWraptext(int row)
        {
            _activeSheet.Row(row).Style.WrapText = true;
        }

        public void SetCellValue(int row, object column, object cellValue, int cellToMerge = 0)
        {
            if (cellValue is Hyperlink hyperlink)
            {
                SetCellHyperlink(row, column, hyperlink.Url);
                BaseSetCellValue(row, column, hyperlink.Value);
                SetCellFontStyle(row, column, FontStyle.Underline);
                SetCellFontColor(row, column, Color.Blue);
            }
            else
            {
                if (cellValue is Cell excelCell)
                    SetCell(row, column, excelCell);
                else if (cellValue is Euro euro)
                    SetCell(row, column, new Cell { Value = euro.Value, Format = "#,##0.00€" });
                else if (cellValue is PercentBase1 percent1)
                    SetCell(row, column, new Cell { Value = percent1.Value, Format = "#0.00%" });
                else if (cellValue is PercentBase100 percent100)
                    SetCell(row, column, new Cell { Value = percent100.Value / 100, Format = "#0.00%" });
                else
                    BaseSetCellValue(row, column, cellValue, cellToMerge);
            }
        }

        private void BaseSetCellValue(int row, object column, object cellValue, int cellToMerge = 0)
        {
            var cell = _activeSheet.Cells[row, column.ToColumnInt()];
            if (cellValue == null)
                cell.Value = null;
            else
                switch (Type.GetTypeCode(cellValue.GetType()))
                {
                    case TypeCode.Int16:
                    case TypeCode.Int32:
                    case TypeCode.Int64:
                        cell.Value = cellValue;
                        cell.Style.Numberformat.Format = "0";
                        break;
                    case TypeCode.Double:
                    case TypeCode.Decimal:
                        cell.Value = cellValue;
                        cell.Style.Numberformat.Format = "0.00";
                        break;
                    case TypeCode.DateTime:
                        cell.Style.Numberformat.Format = "dd/mm/yyyy";
                        cell.Value = ((DateTime)cellValue).ToOADate();
                        break;
                    default:
                        cell.Value = cellValue;
                        break;
                }

            if (cellToMerge > 1)
                _activeSheet.Cells[row, column.ToColumnInt(), row, column.ToColumnInt() + cellToMerge - 1].Merge = true;
        }

        private void SetCellValue(object rowCell, int index, int cellToMerge = 0)
        {
            SetCellValue(_quickFill[_activeSheet.Name].CurrentRow, index + 1, rowCell, cellToMerge);
        }

        public void SetCellComment(int row, object column, string commentValue)
        {
            var cell = _activeSheet.Cells[row, column.ToColumnInt()];
            if (cell.Comment == null)
            {
                cell.AddComment(commentValue, "AssurWare"); //Without author, comment is invalid
            }
            else
            {
                cell.Comment.Text = commentValue;
                cell.Comment.Author = "AssurWare";
            }
        }

        public void SetCell(int row, object column, Cell cellData)
        {
            var cell = _activeSheet.Cells[row, column.ToColumnInt()];
            if (!string.IsNullOrEmpty(cellData.Formula))
                cell.Formula = cellData.Formula;
            else
                cell.Value = cellData.Value;
            if (cellData.BackgroundColor.HasValue)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(cellData.BackgroundColor.Value);
            }
            if (cellData.FontColor.HasValue)
                cell.Style.Font.Color.SetColor(cellData.FontColor.Value);
            if (cellData.FontStyle.HasValue)
            {
                cell.Style.Font.Bold = (cellData.FontStyle.Value & FontStyle.Bold) == FontStyle.Bold;
                cell.Style.Font.Italic = (cellData.FontStyle.Value & FontStyle.Italic) == FontStyle.Italic;
                cell.Style.Font.UnderLine = (cellData.FontStyle.Value & FontStyle.Underline) == FontStyle.Underline;
                cell.Style.Font.Strike = (cellData.FontStyle.Value & FontStyle.Strikeout) == FontStyle.Strikeout;
            }
            if (!string.IsNullOrEmpty(cellData.Format))
                cell.Style.Numberformat.Format = cellData.Format;
        }

        public void SetCellBackground(int row, object column, Color color)
        {
            var cell = _activeSheet.Cells[row, column.ToColumnInt()];
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(color);
        }

        public void SetCellBackgroundFromImportance(int row, object column, bool important)
        {
            this.SetCellBackground(row, column, important ? _red : _yellow);
        }

        public void SetCellFontStyle(int row, object column, FontStyle style)
        {
            var cell = _activeSheet.Cells[row, column.ToColumnInt()];
            cell.Style.Font.Bold = (style & FontStyle.Bold) == FontStyle.Bold;
            cell.Style.Font.Italic = (style & FontStyle.Italic) == FontStyle.Italic;
            cell.Style.Font.UnderLine = (style & FontStyle.Underline) == FontStyle.Underline;
            cell.Style.Font.Strike = (style & FontStyle.Strikeout) == FontStyle.Strikeout;
        }

        public void SetCellFontColor(int row, object column, Color color)
        {
            var cell = _activeSheet.Cells[row, column.ToColumnInt()];
            cell.Style.Font.Color.SetColor(color);
        }

        public void SetCellHyperlink(int row, object column, string hyperLink)
        {
            var cell = _activeSheet.Cells[row, column.ToColumnInt()];
            cell.Hyperlink = new Uri(hyperLink);
        }

        public object GetCellValue(int row, object column)
        {
            var value = Value(row, column);
            return value;
        }

        public T GetCellValue<T>(int row, object column)
        {
            var value = Value<T>(row, column);
            return value;
        }

        public string GetCellValueAsString(int row, object column)
        {
            var value = Value(row, column);

            return value?.ToString() ?? string.Empty;
        }

        public string GetCellValueAsStringTrim(int row, object column)
        {
            var value = Value(row, column);

            return value?.ToString().Trim() ?? string.Empty;
        }

        public string GetCellComment(int row, object column)
        {
            var value = _activeSheet.Cells[row, column.ToColumnInt()].Comment.Text;
            return value;
        }

        public Color GetCellBackgroundColor(int row, object column)
        {
            var cell = _activeSheet.Cells[row, column.ToColumnInt()];
            var backgroundColor = cell.Style.Fill.BackgroundColor;
            return string.IsNullOrEmpty(backgroundColor.Rgb) ? ColorTranslator.FromHtml("#FFFFFFFF") : ColorTranslator.FromHtml("#" + backgroundColor.Rgb);
        }

        public Cell GetCell(int row, object column)
        {
            var excelCell = new Cell();
            var cell = _activeSheet.Cells[row, column.ToColumnInt()];
            excelCell.Value = Value(row, column);
            excelCell.Format = cell.Style.Numberformat.Format;
            excelCell.Formula = cell.Formula;
            excelCell.Hyperlink = cell.Hyperlink;
            return excelCell;
        }

        public int ControlDate(DateTime? date, int yearMin, int yearMax, bool emptyOk, int row = 0, int column = 0, bool important = false)
        {
            var isValid = !(date == null && !emptyOk);
            if (isValid && date != null)
                isValid = date.Value.Year >= yearMin && date.Value.Year <= yearMax;
            if (!isValid && row > 0)
                SetCellBackgroundFromImportance(row, column, important);
            return isValid ? 0 : 1;
        }

        public int ControlStringMinLength(string item, int minLen, bool emptyOk, int row = 0, int column = 0, bool important = false)
        {
            var isValid = !(string.IsNullOrEmpty(item) && !emptyOk);
            if (isValid && !string.IsNullOrEmpty(item))
                isValid = item.Length >= minLen;
            if (!isValid && row > 0)
                SetCellBackgroundFromImportance(row, column, important);
            return isValid ? 0 : 1;
        }

        public int ControlNumberLength(string number, int minLength, int maxLength, bool emptyOk, int row = 0, int column = 0, bool important = false)
        {
            if (string.IsNullOrEmpty(number) && !emptyOk && row > 0)
            {
                SetCellBackgroundFromImportance(row, column, important);
                return 1;
            }
            var isValid = number?.Length >= minLength && number.Length <= maxLength;
            if (!isValid && row > 0)
                SetCellBackgroundFromImportance(row, column, important);
            return isValid ? 0 : 1;
        }

        public void SetColumnAutoWidth(object column)
        {
            _activeSheet.Column(column.ToColumnInt()).AutoFit();
        }

        public void SetColumnWidth(object column, double width)
        {
            _activeSheet.Column(column.ToColumnInt()).Width = width;
        }

        public void SetColumnHideShow(object column, bool hidden)
        {
            _activeSheet.Column(column.ToColumnInt()).Hidden = hidden;
        }

        public void SetColumnHideShow(string columnRange, bool hidden)
        {
            _activeSheet.Cells[columnRange].Style.Hidden = hidden;
        }

        public void SetColumnFormat(object column, string style)
        {
            _activeSheet.Column(column.ToColumnInt()).Style.Numberformat.Format = style;
        }

        public List<string> GetHeaderLabels(int headerRow, object startColumn)
        {
            var column = startColumn.ToColumnInt();
            var list = new List<string>();
            while (true)
            {
                var cellValue = _activeSheet.GetValue<string>(headerRow, column);
                if (cellValue == null)
                    break;
                list.Add(cellValue);
                column++;
            }
            return list;
        }

        public List<string> GetHeaderComments(int headerRow, object startColumn)
        {
            var column = startColumn.ToColumnInt();
            var list = new List<string>();
            while (true)
            {
                var cellComment = _activeSheet.Cells[headerRow, column].Comment?.Text;
                if (cellComment == null)
                    break;
                list.Add(cellComment);
                column++;
            }
            return list;
        }

        public void RemoveHeaderComments(int headerRow, object startColumn)
        {
            var column = startColumn.ToColumnInt();
            while (true)
            {
                var cellComment = _activeSheet.Cells[headerRow, column].Comment?.Text;
                if (cellComment == null)
                    break;
                _activeSheet.Comments.Remove(_activeSheet.Cells[headerRow, column].Comment);
                column++;
            }
        }

        public string TableToJson(List<string> properties, int tableStartRow, object tableStartColumn)
        {
            var isPropertyInteger = new bool[properties.Count];

            //Convention: if '*' is the first character of the property, force the type to 'int' (excel only knows doubles)
            for (var index = 0; index < properties.Count; index++)
            {
                var property = properties[index];
                if (property.StartsWith("*"))
                {
                    properties[index] = property.Remove(0, 1);
                    isPropertyInteger[index] = true;
                }
            }

            var row = tableStartRow;
            var columnStart = tableStartColumn.ToColumnInt();
            var column = columnStart;
            var rowList = new List<JObject>();
            while (true)
            {
                var cell = _activeSheet.Cells[row, column];
                if (cell.Value == null)
                    break;
                var jObjectRow = new JObject();
                foreach (var property in properties)
                {
                    if (property.ToLower() != "empty")
                    {
                        var value = cell.Value;
                        if (value != null)
                            if (isPropertyInteger[column - columnStart])
                                jObjectRow.Add(property, Convert.ToInt32((double)value));
                            else
                                switch (Type.GetTypeCode(value.GetType()))
                                {
                                    case TypeCode.Double:
                                        jObjectRow.Add(property, (double)value);
                                        break;
                                    case TypeCode.DateTime:
                                        jObjectRow.Add(property, (DateTime)value);
                                        break;
                                    default:
                                        jObjectRow.Add(property, value.ToString());
                                        break;
                                }
                    }
                    column++;
                    cell = _activeSheet.Cells[row, column];
                }
                rowList.Add(jObjectRow);
                row++;
                column = columnStart;
            }
            var final = JsonConvert.SerializeObject(rowList);
            return final;
        }

        public void PropertiesToComments<T>(int tableStartRow, object tableStartColumn) where T : class, new()
        {
            var properties = from property in typeof(T).GetProperties()
                             let orderAttribute = property.GetCustomAttributes(typeof(ExcelGenerationOrderAttribute), false).SingleOrDefault() as ExcelGenerationOrderAttribute
                             orderby orderAttribute.Order
                             select property;
            var column = tableStartColumn.ToColumnInt();
            foreach (var propertyInfo in properties)
            {
                var name = propertyInfo.Name;
                var propertyType = propertyInfo.PropertyType;
                if (propertyType == typeof(int) || propertyType == typeof(int?))
                    name = "*" + name;
                SetCellComment(tableStartRow, column, name.ToLower().StartsWith("empty") ? "Empty" : name);
                column++;
            }
        }

        public Dictionary<string, int> MapPropertiesToColumns(List<string> properties, object columnStart)
        {
            var map = new Dictionary<string, int>();
            var column = columnStart.ToColumnInt();
            foreach (var property in properties)
            {
                if (property.ToLower() != "empty")
                    map[property] = column;
                column++;
            }
            return map;
        }

        public DateTime FromSpreadsheetDate(double serialDate)
        {
            if (serialDate > 59)
                serialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(serialDate);
        }

        public void Calculate(bool ignoreConvert = false)
        {
            if (UseAsPoseCalc && !ignoreConvert)
                using (var asPoseWorkbook = new Workbook(new MemoryStream(SaveToBytes())))
                {
                    var activeSheetName = _activeSheet.Name;
                    asPoseWorkbook.CalculateFormula();
                    var memoryStream = new MemoryStream();
                    asPoseWorkbook.Save(memoryStream, SaveFormat.Auto);
                    var asPoseExcelBytes = memoryStream.ToArray();
                    _package = new ExcelPackage(new MemoryStream(asPoseExcelBytes));
                    //DeleteWorksheet("Evaluation Warning");
                    SelectWorksheet(activeSheetName);
                }
            else
                _package.Workbook.Calculate();
        }

        public object SearchLabelAndGetValue(string label, object worksheetsNames = null, int rowOffsetForValue = 0, int columnOffsetForValue = 1, string cellProximity = null, int proximityRadius = 1)
        {
            var worksheetNameList = WorksheetNameList(worksheetsNames);

            foreach (var worksheetName in worksheetNameList)
            {
                var workSheet = _package.Workbook.Worksheets[worksheetName];
                var start = workSheet.Dimension.Start;
                var end = workSheet.Dimension.End;
                for (var row = start.Row; row <= end.Row; row++)
                    for (var column = start.Column; column <= end.Column; column++)
                    {
                        var s = workSheet.GetValue(row, column) as string;
                        if (s != null && label == s)
                        {
                            if (string.IsNullOrEmpty(cellProximity))
                                return workSheet.GetValue(row + rowOffsetForValue, column + columnOffsetForValue);
                            var rowTarget = _activeSheet.Cells[cellProximity].Start.Row;
                            var columnTarget = _activeSheet.Cells[cellProximity].Start.Column;
                            if (row.Between(Math.Max(rowTarget - proximityRadius, 1), rowTarget + proximityRadius) && column.Between(Math.Max(columnTarget - proximityRadius, 1), columnTarget + proximityRadius))
                                return workSheet.GetValue(row + rowOffsetForValue, column + columnOffsetForValue);
                        }
                    }
            }
            throw new SpreadsheetGenerationExceptions.CellNotFoundException($"Cell '{cellProximity}' not found");
        }

        public void SearchLabelAndSetValue(string label, object value, object worksheetsNames = null, int rowOffsetForValue = 0, int columnOffsetForValue = 1, string cellProximity = null, int proximityRadius = 1)
        {
            var worksheetNameList = WorksheetNameList(worksheetsNames);

            foreach (var worksheetName in worksheetNameList)
            {
                var workSheet = _package.Workbook.Worksheets[worksheetName];
                var start = workSheet.Dimension.Start;
                var end = workSheet.Dimension.End;
                for (var row = start.Row; row <= end.Row; row++)
                    for (var column = start.Column; column <= end.Column; column++)
                    {
                        var s = workSheet.GetValue(row, column) as string;
                        if (s == null || label != s)
                            continue;
                        if (string.IsNullOrEmpty(cellProximity))
                        {
                            workSheet.SetValue(row + rowOffsetForValue, column + columnOffsetForValue, value);
                            _package.Workbook.Calculate();
                            return;
                        }
                        var rowTarget = _activeSheet.Cells[cellProximity].Start.Row;
                        var columnTarget = _activeSheet.Cells[cellProximity].Start.Column;
                        if (row.Between(Math.Max(rowTarget - proximityRadius, 1), rowTarget + proximityRadius) && column.Between(Math.Max(columnTarget - proximityRadius, 1), columnTarget + proximityRadius))
                        {
                            workSheet.SetValue(row + rowOffsetForValue, column + columnOffsetForValue, value);
                            _package.Workbook.Calculate();
                            return;
                        }
                    }
            }
            throw new SpreadsheetGenerationExceptions.CellNotFoundException($"Cell '{cellProximity}' not found");
        }

        private static string ToDelimitedString(List<string> list, string delimiter = ":", bool insertSpaces = false, string qualifier = "", bool duplicateTicksForSql = false)
        {
            var result = new StringBuilder();
            for (var i = 0; i < list.Count; i++)
            {
                var initialStr = duplicateTicksForSql ? list[i].Replace("'", "''") : list[i];
                result.Append((qualifier == string.Empty) ? initialStr : string.Format("{1}{0}{1}", initialStr, qualifier));

                if (i < list.Count - 1)
                {
                    result.Append(delimiter);
                    if (insertSpaces)
                        result.Append(' ');
                }
            }
            return result.ToString();
        }

        private List<string> WorksheetNameList(object worksheetsNames)
        {
            var worksheetNameList = new List<string>();
            if (worksheetsNames == null)
                worksheetNameList.AddRange(_package.Workbook.Worksheets.Select(excelWorksheet => excelWorksheet.Name));
            else if (worksheetsNames is string)
                worksheetNameList.Add((string)worksheetsNames == string.Empty ? _activeSheet.Name : (string)worksheetsNames);
            else if (worksheetsNames is List<string>)
                worksheetNameList = (List<string>)worksheetsNames;
            else if (worksheetsNames is int)
                worksheetNameList.Add(_package.Workbook.Worksheets[(int)worksheetsNames].Name);
            else
                throw new SpreadsheetGenerationExceptions.UnknownTypeException("Variable should be of type string or List<string>");
            return worksheetNameList;
        }

        public DateTime FromExcelDate(string dateTime)
        {
            DateTime returnedDate;
            if (DateTime.TryParse(dateTime, out returnedDate)) return returnedDate;
            var serialDate = double.Parse(dateTime);
            if (serialDate > 59)
                serialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(serialDate);
        }

        private object Value(int row, object column)
        {
            return _activeSheet.GetValue(row, column.ToColumnInt());
        }

        private T Value<T>(int row, object column)
        {
            if (typeof(string) == typeof(T))
            {
                // https://stackoverflow.com/questions/1546113/double-to-string-conversion-without-scientific-notation/33697376#33697376
                // Issue with double converting to string if we use ExcelGeneration.Value(x, y).ToString(); using ExcelGeneration.Value<string>(x, y) fixed the issue.
                // See GetContractNumber_test tests for more info
                var value = _activeSheet.Cells[row, column.ToColumnInt()].Value;
                if (value is double valueDouble)
                {
                    return (T)(object)valueDouble.ToString(DoubleConversionFormat);
                }
            }

            return _activeSheet.GetValue<T>(row, column.ToColumnInt());
        }

        private void QuickFilterAndFreeze(string worksheetName, bool freezePanels = true)
        {
            var activeSheet = _activeSheet;
            SelectWorksheet(worksheetName);
            for (var index = 1; index < _quickFill[worksheetName].LastColumnExtend; index++)
                if (!_activeSheet.Column(index).Hidden)
                    SetColumnAutoWidth(index);
            SetWorksheetAutofilter();
            if (freezePanels)
                FreezeWorksheetPanes(2, 1);
            _activeSheet = activeSheet;
        }

        public int GetTotalExcelRows(int startRow, int columnCheck = 1)
        {
            var total = 0;
            while (true)
            {
                var cell = GetCellValue(startRow, columnCheck);
                if (cell == null)
                    break;
                total++;
                startRow++;
            }
            return total;
        }


        #region IDisposable Support
        private bool disposed; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
            {
                return;
            }

            if (disposing)
            {
                // dispose managed state (managed objects) and set variable to null
                _package?.Dispose();
                _package = null;
            }

            disposed = true;
        }

        // Do not change this code. Put cleanup code in Dispose(bool disposing) above.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        #endregion
    }

    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ExcelGenerationOrderAttribute : Attribute
    {
        public ExcelGenerationOrderAttribute([CallerLineNumber] int order = 0)
        {
            Order = order;
        }

        public int Order { get; }
    }


}