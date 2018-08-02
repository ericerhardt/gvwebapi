using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class CyclePeriodScheduleMap : ClassMap<CyclePeriodSchedulesEntity>
    {
        public CyclePeriodScheduleMap()
        {
            Table("CyclePeriodSchedules");

            Id(x => x.CyclePeriodScheduleId).GeneratedBy.Identity();

            Map(x => x.InstancesInvoiced);

            References(x => x.CyclePeriod, "CyclePeriodId").Cascade.None();
            References(x => x.Schedule, "ScheduleId").Cascade.None();
        }
    }
}