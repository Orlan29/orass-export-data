using ExportOrass.BusinessLogic.Interfaces;
using ExportOrass.BusinessLogic.Services;
using ExportOrass.DataAccess.Models;
using InfiSoftware.Common.DataAccess.SpreadsheetGeneration;
using InfiSoftware.Common.DataAccess.SpreadsheetGeneration.Excel;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.Configure<ExportOrassDatabaseSettings>(
    builder.Configuration.GetSection("ExportOrassDatabase"));

builder.Services.AddScoped<IExportData, ExportDataService>();
builder.Services.AddScoped<ISpreadsheetGeneration, ExcelGeneration>();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
