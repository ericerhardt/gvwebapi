using FluentNHibernate.Mapping;
using GV.CoFreedomDomain.Entities;

namespace GV.CoFreedomDomain.Mappings
{
    public class IcModelMap : ClassMap<IcModelEntity>
    {
        public IcModelMap()
        {
            Table("ICModels");

            Id(x => x.ModelId).GeneratedBy.Identity();

            Map(x => x.Model);
        }
    }
}