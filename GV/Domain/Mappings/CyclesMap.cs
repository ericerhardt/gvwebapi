using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class CyclesMap : ClassMap<CyclesEntity>
    {
        public CyclesMap()
        {
            Table("Cycles");

            Id(x => x.CycleId).GeneratedBy.Identity();

            Map(x => x.StartDate);
            Map(x => x.CustomerId);
            Map(x => x.InActive);
            Map(x => x.IsDeleted);
            Map(x => x.CreatedDateTime);
            Map(x => x.ModifiedDateTime);
            Map(x => x.InvisibleToClient);
            Map(x => x.EndDate);
            Map(x => x.IsReconciled);
            Map(x => x.ReconcileAdj);

            HasMany(x => x.CyclePeriods)
                .Cascade.AllDeleteOrphan()
                .KeyColumns.Add("CycleId")
                .Access.CamelCaseField(Prefix.Underscore);

            HasMany(x => x.ReconServices)
                .Cascade.AllDeleteOrphan()
                .KeyColumns.Add("CycleId")
                .Access.CamelCaseField(Prefix.Underscore);
        }
    }
}