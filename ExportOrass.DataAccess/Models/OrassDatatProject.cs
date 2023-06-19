using Fingers10.ExcelExport.Attributes;

namespace ExportOrass.DataAccess.Models
{
    public class OrassDatatProject
    {
        [IncludeInReport]
        public string Police { get; set; } = string.Empty;
        [IncludeInReport]
        public string Avenant { get; set; } = string.Empty;
        [IncludeInReport]
        public DateTime Emission { get; set; }
        [IncludeInReport]
        public DateTime DateCompte { get; set; }
        [IncludeInReport]
        public DateTime Effet { get; set; }
        [IncludeInReport]
        public DateTime Expiration { get; set; }
        [IncludeInReport]
        public DateTime Echeance { get; set; }
        [IncludeInReport]
        public string Duree { get; set; } = string.Empty;
        [IncludeInReport]
        public uint Cat { get; set; }
        [IncludeInReport]
        public string Mouvement { get; set; } = string.Empty;
        [IncludeInReport]
        public string Assure { get; set; } = string.Empty;
        [IncludeInReport]
        public string Immat { get; set; } = string.Empty;
        [IncludeInReport]
        public uint Places { get; set; }
        [IncludeInReport]
        public uint PuissanceAd { get; set; }
        [IncludeInReport]
        public string Genre { get; set; } = string.Empty;
        [IncludeInReport]
        public DateTime DateImmat { get; set; }
        [IncludeInReport]
        public string Conducteur { get; set; } = string.Empty;
        [IncludeInReport]
        public string NPermis { get; set; } = string.Empty;
        [IncludeInReport]
        public DateTime DateNaissance { get; set; }
        [IncludeInReport]
        public int Van { get; set; }
        [IncludeInReport]
        public int Vv { get; set; }
        [IncludeInReport]
        public string Marque { get; set; } = string.Empty;
        [IncludeInReport]
        public uint Barem { get; set; }
        [IncludeInReport]
        public uint Cie { get; set; }
        [IncludeInReport]
        public string NomClient { get; set; } = string.Empty;
        [IncludeInReport]
        public string Adresse { get; set;} = string.Empty;
        [IncludeInReport]
        public string Titre { get; set; } = string.Empty;
        [IncludeInReport]
        public string NClient { get; set; } = string.Empty;
        [IncludeInReport]
        public uint TypeContrat { get; set; }
        [IncludeInReport]
        public string Profession { get; set; } = string.Empty;
        [IncludeInReport]
        public string Rc { get; set; } = null!;
        [IncludeInReport]
        public int TotalPn { get; set; }
        [IncludeInReport]
        public string Prefix { get; set; } = string.Empty;
        [IncludeInReport]
        public string NAttestation { get; set; } = string.Empty;
        [IncludeInReport]
        public string CarteRose { get; set; } = string.Empty;
        [IncludeInReport]
        public string CodeInte { get; set; } = string.Empty;
        [IncludeInReport]
        public string NomIntermediaire { get; set; } = string.Empty;
    }
}
