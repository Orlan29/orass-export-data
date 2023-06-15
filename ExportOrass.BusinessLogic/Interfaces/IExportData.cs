﻿using ExportOrass.BusinessLogic.Services;
using ExportOrass.DataAccess.Models;
using Fingers10.ExcelExport.ActionResults;

namespace ExportOrass.BusinessLogic.Interfaces
{
    public interface IExportData
    {
        public ExcelResult<User> ExportDataToCSV(CancellationToken cancellationToken);
        public Task<IEnumerable<OrassData>> GetOrassDatasAsync(string startedDate, string endedDate, CancellationToken cancellationToken);
    }
}