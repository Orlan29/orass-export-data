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
        [BsonRepresentation(BsonType.ObjectId)]
        public string IntermediaryId { get; set; } = string.Empty;
        [BsonRepresentation(BsonType.ObjectId)]
        public string PrincipalInsuredId { get; set; } = string.Empty;
        public DateTime EffectDate { get; set; }
        public IEnumerable<VehicleRef>? Vehicles { get; set; } = null!;
        public int Step { get; set; }
        public int OperationType { get; set; }
        public IEnumerable<int> Periods { get; set; } = null!;
    }

    [BsonIgnoreExtraElements]
    public class VehicleRef
    {
        public DataRef Data { get; set; } = null!;
        public ProductRef? Product { get; set; } = null!;
        public FreeCombinationRef? FreeCombination { get; set; } = null!;
    }

    [BsonIgnoreExtraElements]
    public class DataRef
    {
        public string Registration { get; set; } = string.Empty;
        public int? Category { get; set; }
        public int FiscalPower { get; set; }
        public int NumberOfSeats { get; set; }
        public string Gender { get; set; } = string.Empty;
        public DateTime FirstRegistration { get; set; }
        public double MarketValue { get; set; }
        public string Manufacturer { get; set; } = string.Empty;
    }

    [BsonIgnoreExtraElements]
    public class ProductRef
    {
        public IEnumerable<ProductsGuaranteesRef> ProductsGuarantees { get; set; } = null!;
    }

    [BsonIgnoreExtraElements]
    public class FreeCombinationRef
    {
        public ProductInfoRef ProductInfo { get; set; } = null!;
    }

    [BsonIgnoreExtraElements]
    public class ProductInfoRef
    {
        public IEnumerable<ProductsGuaranteesRef> ProductsGuarantees { get; set; } = null!;
    }

    [BsonIgnoreExtraElements]
    public class ProductsGuaranteesRef
    {
        public string Title { get; set; } = string.Empty;
        [BsonElement("IssuedPrice")]
        public IEnumerable<IssuedPriceRef> IssuedPrices { get; set; } = null!;
    }

    [BsonIgnoreExtraElements]
    public class IssuedPriceRef
    {
        public int NbMonths { get; set; }
        public decimal Price { get; set; }
    }
}
