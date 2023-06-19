using ExportOrass.BusinessLogic.Interfaces;
using ExportOrass.DataAccess.Models;
using Fingers10.ExcelExport.ActionResults;
using Fingers10.ExcelExport.Attributes;
using Microsoft.Extensions.Options;
using MongoDB.Driver;
using System.Text.Json.Serialization;

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

        public async Task<ExcelResult<OrassDatatProject>> ExportDataToCSVAsync(string startDate, string endDate, CancellationToken cancellationToken)
        {
            IEnumerable<OrassData> datas = await GetOrassDatasAsync(startDate, endDate, cancellationToken);
            List<OrassDatatProject> orassDatatProjects = new();

            foreach (OrassData data in datas)
            {
                orassDatatProjects.Add(new OrassDatatProject
                {
                    Police = data.Contrat.PolicyNumber,
                    Avenant = "A",
                    Emission = data.Contrat.CreatedAt,
                    DateCompte = data.Contrat.CreatedAt,
                    Effet = data.Contrat.EffectDate,
                    Expiration = data.Contrat.DueDate,
                    Echeance = data.Contrat.DueDate,
                    Duree = "A",
                    Cat = data.Quotation.Vehicles.First().Data.Category,
                    Mouvement = "A",
                    Assure = data.Client.FirstName + " " + data.Client.LastName,
                    Immat = data.Quotation.Vehicles.First().Data.Registration,
                    Places = data.Quotation.Vehicles.First().Data.NumberOfSeats,
                    PuissanceAd = data.Quotation.Vehicles.First().FiscalPower,
                    Genre = data.Quotation.Vehicles.First().Data.Gender,
                    DateImmat = data.Quotation.Vehicles.First().Data.FirstRegistration,
                    Conducteur = data.Client.FirstName + " " + data.Client.LastName,
                    DateNaissance = data.Client.BirthDate,
                    NPermis = data.Client.DriverLicenseCategory,
                    Van = data.Quotation.Vehicles.First().Data.MarketValue,
                    Vv = data.Quotation.Vehicles.First().Data.MarketValue,
                    Marque = data.Quotation.Vehicles.First().Data.Manufacturer,
                    Barem = 1,
                    Cie = 0,
                    NomClient = data.Client.FirstName + " " + data.Client.LastName,
                    Adresse = data.Client.Adress,
                    Titre = data.Client.Civility,
                    NClient = data.Client.SignatureId,
                    TypeContrat = 1,
                    Profession = data.Client.Occupation,
                    Rc = data.Quotation.Vehicles.First().FreeCombination.ProductInfo.Code,
                    TotalPn = 0,
                    Prefix = "A",
                    NAttestation = data.CertificateSetting.CertificatesInUse.First().Registration,
                    CarteRose = "A",
                    CodeInte = data.Intermediary.AdministrativeRegistration,
                    NomIntermediaire = data.Intermediary.CorporateName
                });
            }

            return new ExcelResult<OrassDatatProject>(orassDatatProjects, "sheet1", "Orass_Data");
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
