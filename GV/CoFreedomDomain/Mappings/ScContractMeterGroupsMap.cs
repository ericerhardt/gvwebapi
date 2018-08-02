using FluentNHibernate.Mapping;
using GV.CoFreedomDomain.Entities;

namespace GV.CoFreedomDomain.Mappings
{
    public class ScContractMeterGroupsMap : ClassMap<ScContractMeterGroupsEntity>
    {
        public ScContractMeterGroupsMap()
        {
            Table("SCContractMeterGroups");

            Id(x => x.ContractMeterGroupId).GeneratedBy.Identity();

            Map(x => x.ContractMeterGroup);
        }
    }
}