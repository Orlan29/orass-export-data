using ExportOrass.BusinessLogic.Services;
using ExportOrass.DataAccess.Models;
using Fingers10.ExcelExport.ActionResults;

namespace ExportOrass.BusinessLogic.Interfaces
{
    public interface IExportData
    {
        public Task<byte[]> ExportOrassDataAsExcel(string startDate, string endDate, CancellationToken cancellationToken);
        public Task<IEnumerable<OrassData>> GetOrassDatasAsync(string startedDate, string endedDate, CancellationToken cancellationToken);
    }
}
