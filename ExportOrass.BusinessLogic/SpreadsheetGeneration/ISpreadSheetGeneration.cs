using InfiSoftware.Common.DataAccess.SpreadsheetGeneration.Model;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace InfiSoftware.Common.DataAccess.SpreadsheetGeneration
{
    public class Cell
    {
        public object Value { get; set; }
        public string Format { get; set; }
        public string Formula { get; set; }
        public Color? BackgroundColor { get; set; }
        public Color? FontColor { get; set; }
        public FontStyle? FontStyle { get; set; }
        public Uri Hyperlink { get; set; }
    }

    public class Hyperlink
    {
        public string Value { get; set; }
        public string Url { get; set; }
    }

    public class Euro
    {
        public Euro() { }
        public Euro(decimal value) => Value = value;
        public decimal Value { get; set; }
    }

    public class PercentBase1
    {
        public PercentBase1() { }
        public PercentBase1(decimal value) => Value = value;
        public decimal Value { get; set; }
    }

    public class PercentBase100
    {
        public PercentBase100() { }
        public PercentBase100(decimal value) => Value = value;
        public decimal Value { get; set; }
    }

    public interface ISpreadsheetGeneration : IDisposable
    {
        /// <summary>
        /// Create a new Excel sheet in the document
        /// </summary>
        /// <param name="worksheetName">  Excel work sheet name </param>
        /// <param name="columns"> Excel column c </param>
        void AddQuickWorksheet(string worksheetName, params object[] columns);

        /// <summary>
        /// Select a sheet of the Excel document
        /// </summary>
        /// <param name="worksheet"> Excel work sheet </param>
        void SelectWorksheet(object worksheet);

        /// <summary>
        /// Remove a sheet from the Excel document
        /// </summary>
        /// <param name="worksheetName"> Worksheet Excel </param>
        void DeleteWorksheet(string worksheetName);

        /// <summary>
        /// Adds a new worksheet to the Excel document
        /// </summary>
        /// <param name="worksheetName"> Excel work sheet name </param>
        void AddWorkSheet(string worksheetName);


        /// <summary>
        /// Adds and selects a worksheet in the Excel document
        /// </summary>
        /// <param name="worksheetName"> Excel work sheet name </param>
        void AddAndSelectWorkSheet(string worksheetName);

        /// <summary>
        /// Check if any Worksheet exist
        /// </summary>
        bool WorksheetExist();

        /// <summary>
        /// Check if worksheet exist in the Excel document
        /// </summary>
        /// <param name="worksheetName"> Excel work sheet name </param>
        bool WorksheetExist(string worksheetName);

        /// <summary>
        /// Inserting data horizontally (to be separated with commas to move to the next column)
        /// </summary>
        /// <param name="rowCells"> Excel row number </param>
        void QuickFillRow(params object[] rowCells);

        /// <summary>
        /// Adds a value in the next column of the Excel document
        /// </summary>
        /// <param name="rowCells">Excel row number</param>
        void QuickFillRowAppend(params object[] rowCells);

        /// <summary>
        /// Save to file
        /// </summary>
        void Save(string password = null);

        /// <summary>
        /// Save to file
        /// </summary>
        /// <param name="filePath">File path</param>
        /// <param name="password">Password within Excel file (optional)</param>
        void SaveToFile(string filePath, string password = null);

        /// <summary>
        /// Save to file
        /// </summary>
        /// <param name="filePath">File path</param>
        /// <param name="password">Password within Excel file (optional)</param>
        void SaveToFileStamp(string filePath, string password = null);

        /// <summary>
        /// Save to file
        /// </summary>
        /// <param name="filePath">File path</param>
        /// <param name="worksheet">Worksheet as value (1 based) or worksheet name</param>
        /// <param name="delimiter">Cell delimiter (, by default)</param>
        /// <param name="hasQuotesTextQualifier"></param>
        void SaveToCsvFile(string filePath, object worksheet, string delimiter = ",", bool hasQuotesTextQualifier = false);

        /// <summary>
        /// Save Excel file to a byte array
        /// </summary>
        /// <param name="password">Optionnal password</param>
        /// <returns>The Excel file as bytes</returns>
        byte[] SaveToBytes(string password = null, bool freezePanels = true);

        /// <summary>
        /// Save to file
        /// </summary>
        /// <param name="worksheet">Worksheet as value (1 based) or worksheet name</param>
        /// <param name="delimiter">Cell delimiter (, by default)</param>
        /// <param name="hasQuotesTextQualifier"></param>
        /// <returns>The Excel file as csv bytes</returns>
        byte[] SaveToCsvBytes(object worksheet, string delimiter = ",", bool hasQuotesTextQualifier = false);

        void SaveToCsvStream(Stream outStream, object worksheet, string delimiter = ",", bool hasQuotesTextQualifier = false);

        void ImportCsvFile(string filePath, string delimiter = ",", bool hasQuotesTextQualifier = false);

        /// <summary>
        /// ???
        /// </summary>
        /// <param name="lastColumn"></param>
        void FinishFormatting(int lastColumn);

        /// <summary>
        /// ???
        /// </summary>
        /// <param name="lastColumn"></param>
        /// <param name="startFilterCell"></param>
        void FinishFormatting(int lastColumn, string startFilterCell);

        /// <summary>
        /// ???
        /// </summary>
        void SetWorksheetAutofilter();

        /// <summary>
        /// ???
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        void FreezeWorksheetPanes(int row, object column);

        /// <summary>
        /// Set Column Auto Width
        /// </summary>
        /// <param name="column">Column</param>
        void SetColumnAutoWidth(object column);

        /// <summary>
        /// Set Column Width
        /// </summary>
        /// <param name="column">Column</param>
        /// <param name="width">New width</param>
        void SetColumnWidth(object column, double width);

        /// <summary>
        /// Set Column Hide Show
        /// </summary>
        /// <param name="column">column</param>
        /// <param name="hidden">hidden column</param>
        void SetColumnHideShow(object column, bool hidden);

        /// <summary>
        /// Define the column to be hidden
        /// </summary>
        /// <param name="columnRange">Column range</param>
        /// <param name="hidden">hidden column</param>
        void SetColumnHideShow(string columnRange, bool hidden);

        /// <summary>
        /// Set Column Format
        /// </summary>
        /// <param name="column">Column</param>
        /// <param name="style">Font Style </param>
        void SetColumnFormat(object column, string style);

        /// <summary>
        /// All row columns will be in text wrap mode
        /// </summary>
        /// <param name="row">row</param>
        void SetRowWraptext(int row);

        /// <summary>
        /// Insert a value in a specific cell of an Excel document
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <param name="cellValue">Excel document cell value</param>
        /// <param name="cellToMerge">Excel document cell to merge</param>
        void SetCellValue(int row, object column, object cellValue, int cellToMerge = 0);

        /// <summary>
        /// Insert a value in a specific cell of an Excel document
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <param name="commentValue">Excel document cell value</param>
        void SetCellComment(int row, object column, string commentValue);

        /// <summary>
        /// Insert a value in a specific cell of an Excel document
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <param name="cellData">Customizing the cell, see the ExcelCell class</param>
        void SetCell(int row, object column, Cell cellData);

        /// <summary>
        /// Change the background color of a cell
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <param name="color">Background color of cell (ARGB) </param>
        void SetCellBackground(int row, object column, Color color);

        /// <summary>
        /// Change the background color of a cell
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <param name="important">Red if important, otherwise, yellow</param>
        void SetCellBackgroundFromImportance(int row, object column, bool important);

        /// <summary>
        /// Change the text style of an Excel cell
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <param name="style">Selected style (bold / italic / underline)</param>
        void SetCellFontStyle(int row, object column, FontStyle style);

        /// <summary>
        /// Change the text color of the cell
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <param name="color">Text color (ARGB)</param>
        void SetCellFontColor(int row, object column, Color color);

        /// <summary>
        /// Create a hyperlink in an Excel cell
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <param name="hyperLink">Hyperlink</param>
        void SetCellHyperlink(int row, object column, string hyperLink);

        /// <summary>
        /// Obtain the value of a specific cell
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <returns></returns>
        object GetCellValue(int row, object column);

        /// <summary>
        /// Obtain the value of a specific cell as a string. Null is converted to string.Empty
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <returns></returns>
        string GetCellValueAsString(int row, object column);

        /// <summary>
        /// Obtain the value of a specific cell as a string. String is trimmed from leading and trailing spaces. Null is converted to string.Empty
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <returns></returns>
        string GetCellValueAsStringTrim(int row, object column);

        /// <summary>
        /// Obtain the value of a specific cell
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <returns></returns>
        string GetCellComment(int row, object column);

        /// <summary>
        /// Obtain the background color of an Excel cell
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <returns></returns>
        Color GetCellBackgroundColor(int row, object column);

        /// <summary>
        /// Get information from a cell
        /// </summary>
        /// <param name="row">Line number of the Excel document</param>
        /// <param name="column">Excel Document Column</param>
        /// <returns></returns>
        Cell GetCell(int row, object column);

        int ControlDate(DateTime? date, int yearMin, int yearMax, bool emptyOk, int row = 0, int column = 0, bool important = false);
        int ControlStringMinLength(string item, int minLen, bool emptyOk, int row = 0, int column = 0, bool important = false);
        int ControlNumberLength(string number, int minLength, int maxLength, bool emptyOk, int row = 0, int column = 0, bool important = false);

        /// <summary>
        /// Get header labels
        /// </summary>
        /// <param name="headerRow"> header row</param>
        /// <param name="startColumn">start column</param>
        /// <returns></returns>
        List<string> GetHeaderLabels(int headerRow, object startColumn);

        /// <summary>
        /// Get header Comments
        /// </summary>
        /// <param name="headerRow">header row</param>
        /// <param name="startColumn"> start column</param>
        /// <returns></returns>
        List<string> GetHeaderComments(int headerRow, object startColumn);

        /// <summary>
        /// Remove Header Comments
        /// </summary>
        /// <param name="headerRow">header row</param>
        /// <param name="startColumn">starting column</param>
        void RemoveHeaderComments(int headerRow, object startColumn);

        /// <summary>
        /// Table to Json
        /// </summary>
        /// <param name="properties">List of properties</param>
        /// <param name="tableStartRow">starting row</param>
        /// <param name="tableStartColumn">starting column</param>
        /// <returns></returns>
        string TableToJson(List<string> properties, int tableStartRow, object tableStartColumn);

        /// <summary>
        /// Properties to Excel Comments
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="tableStartRow">starting row</param>
        /// <param name="tableStartColumn">starting column</param>
        void PropertiesToComments<T>(int tableStartRow, object tableStartColumn) where T : class, new();

        /// <summary>
        /// ????
        /// </summary>
        /// <param name="properties">list of properties</param>
        /// <param name="columnStart">starting column</param>
        /// <returns></returns>
        Dictionary<string, int> MapPropertiesToColumns(List<string> properties, object columnStart);

        /// <summary>
        /// ???
        /// </summary>
        /// <param name="serialDate"></param>
        /// <returns></returns>
        DateTime FromSpreadsheetDate(double serialDate);

        /// <summary>
        /// Calculate method
        /// <param name="ignoreConvert">experimental performance improvement</param>
        /// </summary>
        void Calculate(bool ignoreConvert = false);

        /// <summary>
        /// Search label and get value
        /// </summary>
        /// <param name="label">label</param>
        /// <param name="worksheetsNames">worksheet name</param>
        /// <param name="rowOffsetForValue">Row offset for value</param>
        /// <param name="columnOffsetForValue">column offset for value</param>
        /// <param name="cellProximity">proximity cell</param>
        /// <param name="proximityRadius">proximity radius</param>
        /// <returns></returns>
        object SearchLabelAndGetValue(string label, object worksheetsNames = null, int rowOffsetForValue = 0, int columnOffsetForValue = 1, string cellProximity = null, int proximityRadius = 1);

        /// <summary>
        /// Search label and set value
        /// </summary>
        /// <param name="label">label</param>
        /// <param name="value">value</param>
        /// <param name="worksheetsNames">worksheet name</param>
        /// <param name="rowOffsetForValue">Row offset for value</param>
        /// <param name="columnOffsetForValue">column offset for value</param>
        /// <param name="cellProximity">proximity cell</param>
        /// <param name="proximityRadius">proximity radius</param>
        void SearchLabelAndSetValue(string label, object value, object worksheetsNames = null, int rowOffsetForValue = 0, int columnOffsetForValue = 1, string cellProximity = null, int proximityRadius = 1);
        
        /// <summary>
        /// Create a new Excel sheet in the document with merged columns
        /// </summary>
        /// <param name="worksheetName">  Excel work sheet name </param>
        /// <param name="columns"> Excel column c </param>
        void AddQuickWorksheetWithMerge(string worksheetName, params FusionCell[] columns);

        /// <summary>
        /// Inserting data horizontally (to be separated with commas to move to the next column) with merged columns
        /// </summary>
        /// <param name="rowCells"> Excel row number </param>
        void QuickFillRowWithMerge(params FusionCell[] rowCells);
    }

}
