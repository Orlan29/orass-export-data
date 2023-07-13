using ExportOrass.BusinessLogic.Interfaces;
using ExportOrass.DataAccess.Models;
using InfiSoftware.Common.DataAccess.SpreadsheetGeneration;
using InfiSoftware.Common.DataAccess.SpreadsheetGeneration.Excel;
using Microsoft.Extensions.Options;
using MongoDB.Driver;
using System.Data;

namespace ExportOrass.BusinessLogic.Services
{
    public class ExportDataService : IExportData
    {
        private readonly IMongoCollection<Client> _clientsCollection;
        private readonly IMongoCollection<Contrat> _contratsCollection;
        private readonly IMongoCollection<CertificateSetting> _certificateSettingsCollection;
        private readonly IMongoCollection<Intermediary> _intermediariesCollection;
        private readonly IMongoCollection<Quotation> _quotationsCollection;
        private readonly ISpreadsheetGeneration _excelGeneration;

        public ExportDataService(IOptions<ExportOrassDatabaseSettings> exportOrassDatabaseSettings, ISpreadsheetGeneration excelGeneration)
        {
            var mongoClient = new MongoClient(exportOrassDatabaseSettings.Value.ConnectionString);
            var mongoDatabase = mongoClient.GetDatabase(exportOrassDatabaseSettings.Value.DatabaseName);

            _clientsCollection = mongoDatabase.GetCollection<Client>("Client");
            _contratsCollection = mongoDatabase.GetCollection<Contrat>("Contract");
            _certificateSettingsCollection = mongoDatabase.GetCollection<CertificateSetting>("CertificateSettings");
            _intermediariesCollection = mongoDatabase.GetCollection<Intermediary>("Intermediary");
            _quotationsCollection = mongoDatabase.GetCollection<Quotation>("Quotation");
            _excelGeneration = excelGeneration;
        }


        public async Task<byte[]> ExportOrassDataAsExcel(string startDate, string endDate, CancellationToken cancellationToken)
        {
            IEnumerable<OrassData> orassDatas = await GetOrassDatasAsync(startDate, endDate, cancellationToken);

            using var excel = new ExcelGeneration("POLICE", "AVENANT", "EMISSION", "Date Comp", "EFFET", "EXPIRATION", "ECHEANCE", "DUREE", "CAT", "MOUVEMENT",
                "ASSURE", "IMMAT", "PLACES", "PUISSANCE AD", "GENRE", "DATE IMMAT", "CONDUCTEUR", "DATE NAISSANCE", "N° PERMIS", "VAN", "VV", "MARQUE", "BAREM",
                "CIE", "NOM CLIENT", "ADRESSE", "TITRE", "N° CLIENT", "TYPE CONTRAT", "PROFESSION", "RC", "TOTALPN", "Prefixe", "N° ATTESTATION", "N° CARTE ROSE", "CODE INTE", "NOM INTERMEDIAIRE");

            foreach (var data in orassDatas)
            {
                var vehicles = data.Quotation.Vehicles;
                string operationType = data.Quotation.OperationType == 0 ? "Emission" : "Renouvellement";
                decimal totalPrice = 0;

                ProductRef? product = data.Quotation.Vehicles?.Where(x => x.Product != null).Select(x => x.Product).FirstOrDefault();
                FreeCombinationRef? freeCombination = data.Quotation.Vehicles?.Where(x => x.FreeCombination != null).Select(x => x.FreeCombination).FirstOrDefault();
                var productsGuarantees = product?.ProductsGuarantees ?? freeCombination?.ProductInfo?.ProductsGuarantees;

                if (productsGuarantees != null)
                    totalPrice = productsGuarantees
                       .SelectMany(x => x.IssuedPrices)
                       .Where(x => x.NbMonths == data.Quotation.Periods.FirstOrDefault())
                       .Sum(x => x.Price);

                if (vehicles != null && vehicles.Any())
                 {
                    excel.QuickFillRow(
                        data.Contrat.PolicyNumber,
                        $"{data.Contrat.ContractDate:dd/MM/yyyy}",
                        $"{data.Contrat.ContractDate:dd/MM/yyyy}",
                        $"{data.Contrat.EffectDate:dd/MM/yyyy}",
                        $"{data.Contrat.DueDate:dd/MM/yyyy}",
                        $"{data.Contrat.DueDate:dd/MM/yyyy}",
                        $"{(data.Contrat.DueDate - data.Contrat.EffectDate).ToString():dd/MM/yyyy}",
                        vehicles?.FirstOrDefault()?.Data?.Category,
                        operationType,
                        data.Client.FirstName + " " + data.Client.LastName,
                        vehicles?.FirstOrDefault()?.Data?.Registration,
                        vehicles?.FirstOrDefault()?.Data?.NumberOfSeats,
                        vehicles?.FirstOrDefault()?.Data?.FiscalPower,
                        vehicles?.FirstOrDefault()?.Data?.Gender,
                        vehicles?.FirstOrDefault()?.Data?.FirstRegistration,
                        data.Client.FirstName + " " + data.Client.LastName,
                        $"{data.Client.BirthDate:dd/MM/yyyy}",
                        data.Client.DriverLicenseCategory,
                        vehicles?.FirstOrDefault()?.Data?.MarketValue,
                        vehicles?.FirstOrDefault()?.Data?.MarketValue,
                        vehicles?.FirstOrDefault()?.Data?.Manufacturer,
                        1,
                        0,
                        data.Client.FirstName + " " + data.Client.LastName,
                        data.Client.Adress,
                        data.Client.Civility,
                        data.Client.SignatureId,
                        1,
                        data.Client.Occupation,
                        vehicles?.FirstOrDefault()?.FreeCombination?.ProductInfo?.ProductsGuarantees?.FirstOrDefault()?.Title,
                        (double)totalPrice,
                        "",
                        data.CertificateSetting.CertificatesInUse.FirstOrDefault()?.Registration,
                        "",
                        data.Intermediary.AdministrativeRegistration,
                        data.Intermediary.CorporateName);
                }
            }

            return excel.SaveToBytes();
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
