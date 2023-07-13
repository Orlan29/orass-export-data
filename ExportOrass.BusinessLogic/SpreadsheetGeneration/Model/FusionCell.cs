namespace InfiSoftware.Common.DataAccess.SpreadsheetGeneration.Model
{
    public class FusionCell
    {
        public object CellValue { get; set; }
        public int CellWidth { get; set; }

        public FusionCell(object cellValue, int cellWidth)
        {
            CellValue = cellValue;
            CellWidth = cellWidth;
        }
    }
}