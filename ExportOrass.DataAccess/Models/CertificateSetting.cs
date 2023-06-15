using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;

namespace ExportOrass.DataAccess.Models
{
    [BsonIgnoreExtraElements]
    public class CertificateSetting
    {
        [BsonId]
        [BsonRepresentation(BsonType.ObjectId)]
        public string Id { get; set; } = string.Empty;
        [BsonElement("IntermediaryId")]
        [BsonRepresentation(BsonType.ObjectId)]
        public string IntermediaryId { get; set; } = string.Empty;
        [BsonElement("CertificatesInUse")]
        public IEnumerable<ProjectCertificatesInUseRef> CertificatesInUse { get; set; } = null!;
    }

    [BsonIgnoreExtraElements]
    public class ProjectCertificatesInUseRef
    {
        [BsonElement("Registration")]
        public string Registration { get; set; } = string.Empty;
    }
}
