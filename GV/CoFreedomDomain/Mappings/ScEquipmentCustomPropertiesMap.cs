using FluentNHibernate.Mapping;
using GV.CoFreedomDomain.Entities;

namespace GV.CoFreedomDomain.Mappings
{
    public class ScEquipmentCustomPropertiesMap : ClassMap<ScEquipmentCustomPropertiesEntity>
    {
        public ScEquipmentCustomPropertiesMap()
        {
            Table("SCEquipmentCustomProperties");

            Id(x => x.EquipmentCustomPropertyId).GeneratedBy.Identity();

            Map(x => x.ShAttributeId);
            Map(x => x.IdVal);
            Map(x => x.TextVal);

            References(x => x.Equipment, "EquipmentId").Cascade.None();
        }
    }
}