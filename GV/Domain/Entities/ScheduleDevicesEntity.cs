using System;

namespace GV.Domain.Entities
{
    public class ScheduleDevicesEntity
    {
        public virtual long ScheduleDeviceID { get; set; }
        public virtual SchedulesEntity Schedule { get; set; }
        public virtual long CustomerID { get; set; }
        public virtual int EquipmentID { get; set; }
        public virtual string EquipmentNumber { get; set; }
        public virtual string SerialNumber { get; set; }
        public virtual string Model { get; set; }
        public virtual string Building { get; set; }
        public virtual string AssetUser { get; set; }
        public virtual string Floor { get; set; }
        public virtual string Department { get; set; }
        public virtual string CostCenter { get; set; }
        public virtual string Address { get; set; }
        public virtual string OwnershipType { get; set; }
        public virtual string ScheduleNumber { get; set; }
        public virtual string IPAddress { get; set; }
        public virtual string Location { get; set; }
        public virtual Nullable<bool> Active { get; set; }
        public virtual Nullable<int> LocationID { get; set; }
        public virtual Nullable<System.DateTime> InstallDate { get; set; }
        public virtual Nullable<int> ContractMeterGroupID { get; set; }
        public virtual string ContractMeterGroup { get; set; }
        public virtual string ModelCategory { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }
        public virtual DateTimeOffset? ModifiedDateTime { get; set; }
        public virtual RemovedStatusEnum? RemovedStatus { get; set; }
        public virtual decimal MonthlyCost { get; set; }
        public virtual DateTimeOffset? RemovalDateTime {get; set; }
        public virtual string Disposition { get; set; }

        public virtual string ChangeKey => $"{EquipmentID}/{EquipmentNumber.Trim().ToLower()}/{SerialNumber.Trim().ToLower()}{Model.Trim().ToLower()}";
    }

    public enum RemovedStatusEnum
    {
        SetForRemoval,
        Removed,
        FormatterReplaced
    }
}