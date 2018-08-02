using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class CycleReconciliationServicesMap : ClassMap<CycleReconciliationServicesEntity>
    {
        public CycleReconciliationServicesMap()
        {
            Table("CycleReconciliationServices");

            Id(x => x.CycleReconciliationServiceId).GeneratedBy.Identity();

            Map(x => x.MeterGroup);
            Map(x => x.Credit);

            References(x => x.Cycle, "CycleId").Cascade.None();
        }
    }
}