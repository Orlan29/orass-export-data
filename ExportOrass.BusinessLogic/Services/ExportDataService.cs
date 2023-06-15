using ExportOrass.BusinessLogic.Interfaces;
using ExportOrass.DataAccess.Models;
using Fingers10.ExcelExport.ActionResults;
using Microsoft.Extensions.Options;
using MongoDB.Driver;

namespace ExportOrass.BusinessLogic.Services
{
    public class User
    {
        public int Id { get; set; }
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

        public ExcelResult<User> ExportDataToCSV(CancellationToken cancellationToken)
        {
            List<User> users = new List<User>
            {
                new User { Id = 1, Username = "DoloresAbernathy" },
                new User { Id = 2, Username = "MaeveMillay" },
                new User { Id = 3, Username = "BernardLowe" },
                new User { Id = 4, Username = "ManInBlack" }
            };

            return new ExcelResult<User>(users, "sheet1", "test");
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
