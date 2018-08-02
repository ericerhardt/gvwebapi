using FluentNHibernate.Mapping;
using GV.CoFreedomDomain.Entities;

namespace GV.CoFreedomDomain.Mappings
{
    public class ScContractDetailsMap : ClassMap<ScContractDetailsEntity>
    {
        public ScContractDetailsMap()
        {
            Table("SCContractDetails");

            Id(x => x.ContractDetailId).GeneratedBy.Identity();

            References(x => x.Contract, "ContractId").Cascade.None();
            References(x => x.Equipment, "EquipmentId").Cascade.None();
        }
    }
}