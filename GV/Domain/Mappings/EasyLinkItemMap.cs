using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class EasyLinkItemMap : ClassMap<EasyLinkItemEntity>
    {
        public EasyLinkItemMap()
        {
            Table("EasyLinkItem");

            Id(x => x.EasyLinkItemId).GeneratedBy.Identity();

            Map(x => x.Child);
            Map(x => x.Email);
            Map(x => x.Fax);
            Map(x => x.DateTime);
            Map(x => x.Description);
            Map(x => x.Duration);
            Map(x => x.Pages);
            Map(x => x.Charge);
            Map(x => x.MessageNumber);

            References(x => x.EasyLink, "EasyLinkId").Cascade.None();
        }
    }
}