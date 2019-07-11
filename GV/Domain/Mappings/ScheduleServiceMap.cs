using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class ScheduleServiceMap : ClassMap<ScheduleServiceEntity>
    {
        public ScheduleServiceMap()
        {
            Table("ScheduleService");

            Id(x => x.ScheduleServiceId).GeneratedBy.Identity();

            Map(x => x.MeterGroup);
            Map(x => x.ContractMeterGroupID);
            Map(x => x.ContractedPages);
            Map(x => x.BaseCpp);
            Map(x => x.OverageCpp);
            Map(x => x.Cost);
            Map(x => x.RemovedFromSchedule);
            Map(x => x.IsDeleted);
            Map(x => x.CreatedDateTime);
            Map(x => x.ModifiedDateTime);
            
            References(x => x.Schedule, "ScheduleId").Cascade.None();
        }
    }
}