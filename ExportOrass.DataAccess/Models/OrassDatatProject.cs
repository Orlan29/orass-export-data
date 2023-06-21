using Fingers10.ExcelExport.Attributes;

namespace ExportOrass.DataAccess.Models
{
    public class OrassDatatProject
    {
        
        public string Police { get; set; } = string.Empty;
        
        public string Avenant { get; set; } = string.Empty;
        
        public DateTime Emission { get; set; }
        
        public DateTime DateCompte { get; set; }
        
        public DateTime Effet { get; set; }
        
        public DateTime Expiration { get; set; }
        
        public DateTime Echeance { get; set; }
        
        public string Duree { get; set; } = string.Empty;
        
        public uint? Cat { get; set; }
        
        public string Mouvement { get; set; } = string.Empty;
        
        public string Assure { get; set; } = string.Empty;
        
        public string Immat { get; set; } = string.Empty;
        
        public uint Places { get; set; }
        
        public uint PuissanceAd { get; set; }
        
        public string Genre { get; set; } = string.Empty;
        
        public DateTime DateImmat { get; set; }
        
        public string Conducteur { get; set; } = string.Empty;
        
        public string NPermis { get; set; } = string.Empty;
        
        public DateTime DateNaissance { get; set; }
        
        public int Van { get; set; }
        
        public int Vv { get; set; }
        
        public string Marque { get; set; } = string.Empty;
        
        public uint Barem { get; set; }
        
        public uint Cie { get; set; }
        
        public string NomClient { get; set; } = string.Empty;
        
        public string Adresse { get; set;} = string.Empty;
        
        public string Titre { get; set; } = string.Empty;
        
        public string NClient { get; set; } = string.Empty;
        
        public uint TypeContrat { get; set; }
        
        public string Profession { get; set; } = string.Empty;
        
        public string Rc { get; set; } = null!;
        
        public int TotalPn { get; set; }
        
        public string Prefix { get; set; } = string.Empty;
        
        public string NAttestation { get; set; } = string.Empty;
        
        public string CarteRose { get; set; } = string.Empty;
        
        public string CodeInte { get; set; } = string.Empty;
        
        public string NomIntermediaire { get; set; } = string.Empty;
    }
}
