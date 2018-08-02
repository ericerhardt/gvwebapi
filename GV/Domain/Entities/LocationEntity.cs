using System;

namespace GV.Domain.Entities
{
    public class LocationEntity
    {
        public LocationEntity(long customerId, string name, int coFreedomLocationId)
        {
            CustomerId = customerId;
            Name = name;
            CoFreedomLocationId = coFreedomLocationId;
            CreatedDateTime = DateTimeOffset.Now;
        }

        public LocationEntity()
        {
        }

        public virtual long LocationId { get; set; }
        public virtual long CustomerId { get; set; }
        public virtual string Name { get; set; }
        public virtual bool IsCorporate { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }
        public virtual DateTimeOffset? ModifiedDateTime { get; set; }
        public virtual bool IsDeleted { get; set; }
        public virtual decimal TaxRate {get; set; }
        public virtual int CoFreedomLocationId { get; set; }
    }
}