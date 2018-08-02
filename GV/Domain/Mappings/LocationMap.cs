using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class LocationMap : ClassMap<LocationEntity>
    {
        public LocationMap()
        {
            Table("Location");
            Schema("dbo");

            Id(x => x.LocationId).GeneratedBy.Identity();

            Map(x => x.CustomerId);
            Map(x => x.Name);
            Map(x => x.IsCorporate);
            Map(x => x.CreatedDateTime);
            Map(x => x.ModifiedDateTime);
            Map(x => x.IsDeleted);
            Map(x => x.TaxRate);
            Map(x => x.CoFreedomLocationId);
        }
    }
}