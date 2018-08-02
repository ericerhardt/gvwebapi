using System;

namespace GV.Domain.Entities
{
    public class DevicesEntity
    {
        public virtual long DeviceId { get; set; }
        public virtual SchedulesEntity Schedule { get; set; }
        public virtual long CustomerId { get; set; }
        public virtual int EquipmentId { get; set; }
        public virtual string EquipmentNumber { get; set; }
        public virtual string SerialNumber { get; set; }
        public virtual string Model { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }
        public virtual DateTimeOffset? ModifiedDateTime { get; set; }
        public virtual RemovedStatusEnum? RemovedStatus { get; set; }
        public virtual decimal MonthlyCost { get; set; }
        public virtual DateTimeOffset? RemovalDateTime {get; set; }
        public virtual string Disposition { get; set; }

        public virtual string ChangeKey => $"{EquipmentId}/{EquipmentNumber.Trim().ToLower()}/{SerialNumber.Trim().ToLower()}{Model.Trim().ToLower()}";
    }

    public enum RemovedStatusEnum
    {
        SetForRemoval,
        Removed,
        FormatterReplaced
    }
}