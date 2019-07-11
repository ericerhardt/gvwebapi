using System;

namespace GV.Domain.Entities
{
    public class ScheduleServiceEntity
    {
        public ScheduleServiceEntity(string meterGroup)
        {
            MeterGroup = meterGroup;
            CreatedDateTime = DateTimeOffset.Now;
        }

        public ScheduleServiceEntity()
        {
        }

        public virtual long ScheduleServiceId { get; set; }
        public virtual SchedulesEntity Schedule { get; set; }
        public virtual string MeterGroup { get; set; }
        public virtual int ContractMeterGroupID { get; set; }
        public virtual int ContractedPages { get; set; }
        public virtual decimal BaseCpp { get; set; }
        public virtual decimal OverageCpp { get; set; }
        public virtual decimal Cost { get; set; }
        public virtual bool RemovedFromSchedule { get; set; }
        public virtual bool IsDeleted { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }
        public virtual DateTimeOffset? ModifiedDateTime { get; set; }

        public virtual void SetCost()
        {
            Cost = ContractedPages * BaseCpp;
        }
    }
}