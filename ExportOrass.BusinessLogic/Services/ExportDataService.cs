using ExportOrass.BusinessLogic.Interfaces;
using ExportOrass.DataAccess.Models;
using Fingers10.ExcelExport.Attributes;
using Microsoft.Extensions.Options;
using MongoDB.Driver;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;

namespace ExportOrass.BusinessLogic.Services
{
    public class User
    {
        [IncludeInReport]
        public int Id { get; set; }
        [IncludeInReport]
        public string Username { get; set; } = string.Empty;
    }

    public class ExportDataService : IExportData
    {
        private readonly IMongoCollection<Client> _clientsCollection;
        private readonly IMongoCollection<Contrat> _contratsCollection;
        private readonly IMongoCollection<CertificateSetting> _certificateSettingsCollection;
        private readonly IMongoCollection<Intermediary> _intermediariesCollection;
        private readonly IMongoCollection<Quotation> _quotationsCollection;

        public ExportDataService(IOptions<ExportOrassDatabaseSettings> exportOrassDatabaseSettings)
        {
            var mongoClient = new MongoClient(exportOrassDatabaseSettings.Value.ConnectionString);
            var mongoDatabase = mongoClient.GetDatabase(exportOrassDatabaseSettings.Value.DatabaseName);

            _clientsCollection = mongoDatabase.GetCollection<Client>("Client");
            _contratsCollection = mongoDatabase.GetCollection<Contrat>("Contract");
            _certificateSettingsCollection = mongoDatabase.GetCollection<CertificateSetting>("CertificateSettings");
            _intermediariesCollection = mongoDatabase.GetCollection<Intermediary>("Intermediary");
            _quotationsCollection = mongoDatabase.GetCollection<Quotation>("Quotation");
        }

        public async Task<byte[]> ExportDataToCSVAsync(string startDate, string endDate, CancellationToken cancellationToken)
        {
            IEnumerable<OrassData> datas = await GetOrassDatasAsync(startDate, endDate, cancellationToken);
            IWorkbook workbook = new XSSFWorkbook();

            var dataFormat = workbook.CreateDataFormat();
            var dataStyle = workbook.CreateCellStyle();
            dataStyle.DataFormat = dataFormat.GetFormat("dd/MM/yyyy");

            ISheet sheet = workbook.CreateSheet("Sheet1");

            int rowNumber = 0;
            IRow row = sheet.CreateRow(rowNumber++);

            InitSheetHeaders(row);

            InitSheetBody(ref sheet, datas);

            MemoryStream ms = new();
            workbook.Write(ms, false);

            byte[] bytes = ms.ToArray();
            ms.Close();

            return bytes;
        }

        private static ICell InitSheetHeaders(IRow row)
        {
            string[] sheetHeaders = { "POLICE", "AVENANT", "EMISSION", "Date Comp", "EFFET", "EXPIRATION", "ECHEANCE", "DUREE", "CAT", "MOUVEMENT", 
                "ASSURE", "IMMAT", "PLACES", "PUISSANCE AD", "GENRE", "DATE IMMAT", "CONDUCTEUR", "DATE NAISSANCE", "N° PERMIS", "VAN", "VV", "MARQUE", "BAREM", "CIE", "NOM CLIENT", "ADRESSE", "TITRE", "N° CLIENT", "TYPE CONTRAT", "PROFESSION", "RC", "TOTALPN", "Prefixe", "N° ATTESTATION", "N° CARTE ROSE", "CODE INTE", "NOM INTERMEDIAIRE" };

            ICell cell = row.CreateCell(0);
            cell.SetCellValue(sheetHeaders[0]);

            for (int i = 1; i < sheetHeaders.Length; i++)
            {
                cell = row.CreateCell(i);
                cell.SetCellValue(sheetHeaders[i]);
            }

            return cell;
        }

        private static void InitSheetBody(ref ISheet sheet, IEnumerable<OrassData> orassDatas)
        {
            int rowNumber = 1;

            foreach (OrassData data in orassDatas)
            {
                var vehicles = data.Quotation.Vehicles;

                if (vehicles != null && vehicles.Any())
                {
                    IRow row = sheet.CreateRow(rowNumber);
                    ICell cell = row.CreateCell(0);
                    cell.SetCellValue(data.Contrat.PolicyNumber);
                    cell = row.CreateCell(1);
                    cell.SetCellValue("");
                    cell = row.CreateCell(2);
                    cell.SetCellValue(data.Contrat.CreatedAt.ToShortDateString());
                    cell = row.CreateCell(3);
                    cell.SetCellValue(data.Contrat.CreatedAt.ToShortDateString());
                    cell = row.CreateCell(4);
                    cell.SetCellValue(data.Contrat.EffectDate.ToShortDateString());
                    cell = row.CreateCell(5);
                    cell.SetCellValue(data.Contrat.DueDate.ToShortDateString());
                    cell = row.CreateCell(6);
                    cell.SetCellValue(data.Contrat.DueDate.ToShortDateString());
                    cell = row.CreateCell(7);
                    cell.SetCellValue((data.Contrat.DueDate - data.Contrat.EffectDate).Days);
                    cell = row.CreateCell(8);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.Data.Category.ToString());
                    cell = row.CreateCell(9);
                    cell.SetCellValue("");
                    cell = row.CreateCell(10);
                    cell.SetCellValue(data.Client.FirstName + " " + data.Client.LastName);
                    cell = row.CreateCell(11);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.Data.Registration);
                    cell = row.CreateCell(12);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.Data.NumberOfSeats.ToString());
                    cell = row.CreateCell(13);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.FiscalPower.ToString());
                    cell = row.CreateCell(14);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.Data.Gender);
                    cell = row.CreateCell(15);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.Data.FirstRegistration.ToShortDateString());
                    cell = row.CreateCell(16);
                    cell.SetCellValue(data.Client.FirstName + " " + data.Client.LastName);
                    cell = row.CreateCell(17);
                    cell.SetCellValue(data.Client.BirthDate.ToShortDateString());
                    cell = row.CreateCell(18);
                    cell.SetCellValue(data.Client.DriverLicenseCategory);
                    cell = row.CreateCell(19);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.Data.MarketValue.ToString());
                    cell = row.CreateCell(20);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.Data.MarketValue.ToString());
                    cell = row.CreateCell(21);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.Data.Manufacturer.ToString());
                    cell = row.CreateCell(22);
                    cell.SetCellValue(1);
                    cell = row.CreateCell(23);
                    cell.SetCellValue(0);
                    cell = row.CreateCell(24);
                    cell.SetCellValue(data.Client.FirstName + " " + data.Client.LastName);
                    cell = row.CreateCell(25);
                    cell.SetCellValue(data.Client.Adress);
                    cell = row.CreateCell(26);
                    cell.SetCellValue(data.Client.Civility);
                    cell = row.CreateCell(27);
                    cell.SetCellValue(data.Client.SignatureId);
                    cell = row.CreateCell(28);
                    cell.SetCellValue(1);
                    cell = row.CreateCell(29);
                    cell.SetCellValue(data.Client.Occupation);
                    cell = row.CreateCell(30);
                    cell.SetCellValue(vehicles?.FirstOrDefault()?.FreeCombination?.ProductInfo?.Code);
                    cell = row.CreateCell(31);
                    cell.SetCellValue(0);
                    cell = row.CreateCell(32);
                    cell.SetCellValue("");
                    cell = row.CreateCell(33);
                    cell.SetCellValue(data.CertificateSetting.CertificatesInUse.FirstOrDefault()?.Registration);
                    cell = row.CreateCell(34);
                    cell.SetCellValue("");
                    cell = row.CreateCell(35);
                    cell.SetCellValue(data.Intermediary.AdministrativeRegistration);
                    cell = row.CreateCell(36);
                    cell.SetCellValue(data.Intermediary.CorporateName);

                    rowNumber++;
                }
            }
        }

        public async Task<IEnumerable<OrassData>> GetOrassDatasAsync(string startDate, string endDate, CancellationToken cancellationToken)
        {

            var quotations = await _quotationsCollection
                 .Find(x =>x.Step == 3 || x.Step == 4)
                 .ToListAsync(cancellationToken);

            var orassData = from quotation in quotations
                                               join intermediary in _intermediariesCollection.AsQueryable() on quotation.IntermediaryId equals intermediary.Id
                                               join certificationSetting in _certificateSettingsCollection.AsQueryable() on intermediary.Id equals certificationSetting.IntermediaryId
                                               join client in _clientsCollection.AsQueryable() on quotation.PrincipalInsuredId equals client.Id
                                               join contract in _contratsCollection.AsQueryable() on client.Id equals contract.ClientId
                                               where quotation.EffectDate >= DateTime.Parse(startDate) && quotation.EffectDate <= DateTime.Parse(endDate)
                                               select new OrassData()
                                               {
                                                   Quotation = quotation,
                                                   Intermediary = intermediary,
                                                   Client = client,
                                                   Contrat = contract,
                                                   CertificateSetting = certificationSetting
                                               };

            return orassData.ToList();
        }
    }
}
