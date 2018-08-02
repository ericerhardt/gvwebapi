using FluentNHibernate.Mapping;
using GV.CoFreedomDomain.Entities;

namespace GV.CoFreedomDomain.Mappings
{
    public class ScEquipmentMap : ClassMap<ScEquipmentEntity>
    {
        public ScEquipmentMap()
        {
            Table("SCEquipments");

            Id(x => x.EquipmentId).GeneratedBy.Identity();

            Map(x => x.EquipmentNumber);
            Map(x => x.CustomerId);
            Map(x => x.Active);
            Map(x => x.SerialNumber);
            
            References(x => x.Location, "LocationId").Cascade.None();
            References(x => x.Model, "ModelId").Cascade.None();

            HasMany(x => x.CustomProperties)
                .PropertyRef("EquipmentId")
                .Cascade.None()
                .KeyColumns.Add("EquipmentId")
                .Access.CamelCaseField(Prefix.Underscore);

            HasMany(x => x.ContractDetails)
                .Cascade.AllDeleteOrphan()
                .KeyColumns.Add("EquipmentId")
                .Access.CamelCaseField(Prefix.Underscore);
        }
    }
}