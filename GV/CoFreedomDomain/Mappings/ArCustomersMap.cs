using FluentNHibernate.Mapping;
using GV.CoFreedomDomain.Entities;

namespace GV.CoFreedomDomain.Mappings
{
    public class ArCustomersMap : ClassMap<ArCustomersEntity>
    {
        public ArCustomersMap()
        {
            Table("ArCustomers");

            Id(x => x.CustomerId).GeneratedBy.Identity();
            Map(x => x.CustomerName);
            Map(x => x.Address);
            Map(x => x.City);
            Map(x => x.State);
            Map(x => x.Zip);
        }
    }
}