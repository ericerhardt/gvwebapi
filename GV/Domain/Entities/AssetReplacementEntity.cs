using System;

namespace GV.Domain.Entities
{
    public class AssetReplacementEntity
    {
        public virtual int ReplacementId { get; set; }
        public virtual int? CustomerId { get; set; }
        public virtual string Location { get; set; }
        public virtual DateTime? ReplacementDate { get; set; }
        public virtual string OldSerialNumber { get; set; }
        public virtual string OldModel { get; set; }
        public virtual string NewSerialNumber { get; set; }
        public virtual string NewModel { get; set; }
        public virtual decimal? ReplacementValue { get; set; }
        public virtual string Comments { get; set; }
        public virtual string ScheduleNumber { get; set; }
    }
}