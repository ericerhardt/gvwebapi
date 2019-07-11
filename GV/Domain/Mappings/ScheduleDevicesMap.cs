using FluentNHibernate.Mapping;
using GV.Domain.Entities;

namespace GV.Domain.Mappings
{
    public class ScheduleDevicesMap : ClassMap<ScheduleDevicesEntity>
    {
        public ScheduleDevicesMap()
        {
            Table("ScheduleDevices");
            Id(x => x.ScheduleDeviceID).GeneratedBy.Identity();
            Map(x => x.CustomerID);
            Map(x => x.EquipmentID);
            Map(x => x.EquipmentNumber);
            Map(x => x.Model);
            Map(x => x.SerialNumber);
            Map(x => x.Active);
            Map(x => x.Address);
            Map(x => x.AssetUser);
            Map(x => x.Building);
            Map(x => x.CostCenter);
            Map(x => x.Department);
            Map(x => x.Floor);
            Map(x => x.IPAddress);
            Map(x => x.ModelCategory);
            Map(x => x.ScheduleNumber);
            Map(x => x.OwnershipType);
            Map(x => x.LocationID);
            Map(x => x.RemovedStatus);
            Map(x => x.MonthlyCost);
            Map(x => x.CreatedDateTime);
            Map(x => x.ModifiedDateTime);
            Map(x => x.RemovalDateTime);
            Map(x => x.Disposition);
            
            References(x => x.Schedule, "ScheduleId")
                .NotFound.Ignore()
                .Cascade.None();
        }
    }
}