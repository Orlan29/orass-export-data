using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace ExportOrass.DataAccess.Models
{
    [BsonIgnoreExtraElements]
    public class Contrat
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = string.Empty;
        [BsonElement("ClientId")]
        [BsonRepresentation(BsonType.ObjectId)]
        public string ClientId { get; set; } = string.Empty;
        [BsonElement("PolicyNumber")]
        public string PolicyNumber { get; set; } = string.Empty;
        [BsonElement("ContractDate")]
        public DateTime ContractDate { get; set; }
        [BsonElement("EffectDate")]
        public DateTime EffectDate { get; set; }
        [BsonElement("DueDate")]
        public DateTime DueDate { get; set; }
    }
}
