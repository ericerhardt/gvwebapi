using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class SchedulesMap : ClassMap<SchedulesEntity>
    {
        public SchedulesMap()
        {
            Table("Schedules");
            Schema("dbo");

            Id(x => x.ScheduleId).GeneratedBy.Identity();

            Map(x => x.CustomerId);
            Map(x => x.Name);
            Map(x => x.EffectiveDateTime);
            Map(x => x.ExpiredDateTime);
            Map(x => x.Term);
            Map(x => x.MonthlyHwCost);
            Map(x => x.MonthlySvcCost);
            Map(x => x.CreatedDateTime);
            Map(x => x.IsDeleted);
            Map(x => x.ModifiedDateTime);

            References(x => x.CoterminousSchedule, "CoterminousScheduleId")
                .NotFound.Ignore()
                .Cascade.None();

            HasMany(x => x.Devices)
                .KeyColumns.Add("ScheduleId")
                .Cascade.AllDeleteOrphan()
                .Access.CamelCaseField(Prefix.Underscore);

            HasMany(x => x.ScheduleServices)
                .KeyColumns.Add("ScheduleId")
                .Cascade.AllDeleteOrphan()
                .Access.CamelCaseField(Prefix.Underscore);
        }
    }
}