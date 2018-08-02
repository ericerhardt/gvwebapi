using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class EasyLinkMap : ClassMap<EasyLinkEntity>
    {
        public EasyLinkMap()
        {
            Table("EasyLink");

            Id(x => x.EasyLinkId).GeneratedBy.Identity();

            Map(x => x.FileName);
            Map(x => x.FileLocation);
            Map(x => x.CreatedDateTime);
            Map(x => x.SavedFileName);
            Map(x => x.NumberOfLines);

            HasMany(x => x.Items)
                .KeyColumns.Add("EasyLInkId")
                .Inverse()
                .Cascade.AllDeleteOrphan()
                .Access.CamelCaseField(Prefix.Underscore);
        }
    }
}