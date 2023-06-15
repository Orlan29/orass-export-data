using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace ExportOrass.DataAccess.Models
{
    [BsonIgnoreExtraElements]
    public class Intermediary
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = string.Empty;
        [BsonElement("CorporateName")]
        public string CorporateName { get; set; } = string.Empty;
        [BsonElement("AdministrativeRegistration")]
        public string AdministrativeRegistration { get; set; } = string.Empty;
    }
}
