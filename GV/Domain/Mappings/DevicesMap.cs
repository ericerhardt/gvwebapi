using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class DevicesMap : ClassMap<DevicesEntity>
    {
        public DevicesMap()
        {
            Table("Devices");

            Id(x => x.DeviceId).GeneratedBy.Identity();

            Map(x => x.Model);
            Map(x => x.CustomerId);
            Map(x => x.SerialNumber);
            Map(x => x.CreatedDateTime);
            Map(x => x.ModifiedDateTime);
            Map(x => x.EquipmentId);
            Map(x => x.EquipmentNumber);
            Map(x => x.RemovedStatus);
            Map(x => x.MonthlyCost);
            Map(x => x.RemovalDateTime);
            Map(x => x.Disposition);
            
            References(x => x.Schedule, "ScheduleId")
                .NotFound.Ignore()
                .Cascade.None();
        }
    }
}