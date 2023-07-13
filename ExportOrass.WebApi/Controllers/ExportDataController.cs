using ExportOrass.BusinessLogic.Interfaces;
using Microsoft.AspNetCore.Mvc;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace ExportOrass.WebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExportDataController : ControllerBase
    {
        private readonly IExportData _exportData;
        public ExportDataController(IExportData exportData)
        {
            _exportData = exportData;
        }

        [HttpGet]
        public async Task<IActionResult> Get(string startedDate,string endedDate, CancellationToken cancellationToken)
        {
            return Ok(await _exportData.GetOrassDatasAsync(startedDate, endedDate, cancellationToken));
        }

        [HttpGet("ExportData")]
        public async Task<IActionResult> ExportData(string startDate, string endDate, CancellationToken cancellationToken)
        {
            var bytes = await _exportData.ExportOrassDataAsExcel(startDate, endDate, cancellationToken);

            return File(bytes, "application/vnd.ms-excel", "Orass_Data.xlsx");
        }
    }
}
