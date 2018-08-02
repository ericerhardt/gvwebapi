using FluentNHibernate.Mapping;
using GV.CoFreedomDomain.Entities;

namespace GV.CoFreedomDomain.Mappings
{
    public class ScContractsMap : ClassMap<ScContractsEntity>
    {
        public ScContractsMap()
        {
            Table("SCContracts");

            Id(x => x.ContractId).GeneratedBy.Identity();

            Map(x => x.CustomerId);

            HasMany(x => x.ContractMeterGroups)
                .Cascade.AllDeleteOrphan()
                .KeyColumns.Add("ContractId")
                .Access.CamelCaseField(Prefix.Underscore);
        }
    }
}