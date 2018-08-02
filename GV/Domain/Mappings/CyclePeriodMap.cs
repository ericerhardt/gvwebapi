using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class CyclePeriodMap : ClassMap<CyclePeriodEntity>
    {
        public CyclePeriodMap()
        {
            Table("CyclePeriods");

            Id(x => x.CyclePeriodId).GeneratedBy.Identity();

            Map(x => x.Period);
            Map(x => x.CreatedDateTime);
            Map(x => x.ModifiedDateTime);
            Map(x => x.IsDeleted);
            Map(x => x.InvoiceNumber);

            References(x => x.Cycle, "CycleId").Cascade.None();

            HasMany(x => x.PeriodSchedules)
                .Cascade.AllDeleteOrphan()
                .KeyColumns.Add("CyclePeriodId")
                .Access.CamelCaseField(Prefix.Underscore);
        }
    }
}