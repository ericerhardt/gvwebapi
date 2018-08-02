using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class EasyLinkChildMatchMap : ClassMap<EasyLinkChildMatchEntity>
    {
        public EasyLinkChildMatchMap()
        {
            Table("EasyLinkChildMatch");

            Id(x => x.EasyLinkChildMatchId).GeneratedBy.Identity();

            Map(x => x.CustomerId);
            Map(x => x.ChildId);
            Map(x => x.IsEasyLinkOnly);
            Map(x => x.CreatedDateTime);
            Map(x => x.ModifiedDateTime);
            Map(x => x.IsDeleted);
        }
    }
}