using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;


namespace ExportOrass.DataAccess.Models
{
    [BsonIgnoreExtraElements]
    public class Quotation
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = string.Empty;
        [BsonElement("IntermediaryId")]
        [BsonRepresentation(BsonType.ObjectId)]
        public string IntermediaryId { get; set; } = string.Empty;
        [BsonElement("PrincipalInsuredId")]
        [BsonRepresentation(BsonType.ObjectId)]
        public string PrincipalInsuredId { get; set; } = string.Empty;
        [BsonElement("EffectDate")]
        public DateTime EffectDate { get; set; }
        [BsonElement("Vehicles")]
        public IEnumerable<ProjectVehicleRef> Vehicles { get; set; } = null!;
        [BsonElement("Step")]
        public uint Step { get; set; }
        [BsonElement("OperationType")]
        public uint OperationType { get; set; }
    }

    [BsonIgnoreExtraElements]
    public class ProjectVehicleRef
    {
        [BsonElement("Data")]
        public ProjectDataRef Data { get; set; } = null!;
        [BsonElement("FiscalPower")]
        public uint FiscalPower { get; set; }
        [BsonElement("FreeCombination")]
        public ProjectFreeCombinationRef FreeCombination { get; set; } = null!;
    }

    [BsonIgnoreExtraElements]
    public class ProjectProductInfoRef
    {
        [BsonElement("Code")]
        public string Code { get; set; } = string.Empty;
    }

    [BsonIgnoreExtraElements]
    public class ProjectFreeCombinationRef
    {
        [BsonElement("ProductInfo")]
        public ProjectProductInfoRef ProductInfo { get; set; } = null!;
    }

    [BsonIgnoreExtraElements]
    public class ProjectDataRef
    {
        [BsonElement("Registration")]
        public string Registration { get; set; } = string.Empty;
        [BsonElement("Category")]
        public uint Category { get; set; }
        [BsonElement("NumberOfSeats")]
        public uint NumberOfSeats { get; set; }
        [BsonElement("Gender")]
        public string Gender { get; set; } = string.Empty;
        [BsonElement("FirstRegistration")]
        public DateTime FirstRegistration { get; set; }
        [BsonElement("MarketValue")]
        public int MarketValue { get; set; }
        [BsonElement("Manufacturer")]
        public string Manufacturer { get; set; } = string.Empty;
    }
}
