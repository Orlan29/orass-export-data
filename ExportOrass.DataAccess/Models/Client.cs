using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace ExportOrass.DataAccess.Models
{
    [BsonIgnoreExtraElements]
    public class Client
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = string.Empty;
        /*[BsonElement("IntermediaryId")]
        [BsonRepresentation(BsonType.ObjectId)]
        public string IntermediaryId { get; set; } = string.Empty;*/
        [BsonElement("LastName")]
        public string LastName { get; set; } = string.Empty;
        [BsonElement("FirstName")]
        public string FirstName { get; set; } = string.Empty;
        [BsonElement("Adress")]
        public string Adress { get; set; } = string.Empty;
        [BsonElement("Civility")]
        public string Civility { get; set; } = string.Empty;
        [BsonElement("BirthDate")]
        public DateTime BirthDate { get; set; }
        [BsonElement("Occupation")]
        public string Occupation { get; set; } = string.Empty;
        [BsonElement("SignatureId")]
        public string SignatureId { get; set; } = string.Empty;
        [BsonElement("DriverLicenseCategory")]
        public string DriverLicenseCategory { get; set; } = string.Empty;
    }
}
