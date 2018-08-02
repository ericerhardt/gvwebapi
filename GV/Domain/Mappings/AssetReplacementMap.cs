using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class AssetReplacementMap  : ClassMap<AssetReplacementEntity>
    {
        public AssetReplacementMap()
        {
            Table("AssetReplacement");

            Id(x => x.ReplacementId).GeneratedBy.Identity();

            Map(x => x.CustomerId);
            Map(x => x.Location);
            Map(x => x.ReplacementDate);
            Map(x => x.OldSerialNumber);
            Map(x => x.OldModel);
            Map(x => x.NewSerialNumber);
            Map(x => x.NewModel);
            Map(x => x.ReplacementValue);
            Map(x => x.Comments);
            Map(x => x.ScheduleNumber);
        }
    }
}