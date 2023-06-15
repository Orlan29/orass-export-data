namespace ExportOrass.DataAccess.Models
{
    public class OrassData
    {
        public Contrat Contrat { get; set; } = null!;
        public Client Client { get; set; } = null!;
        public Quotation Quotation { get; set; } = null!;
        public CertificateSetting CertificateSetting { get; set; } = null!;
        public Intermediary Intermediary { get; set; } = null!;
    }
}
