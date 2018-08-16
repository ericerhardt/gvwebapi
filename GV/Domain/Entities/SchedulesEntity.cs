using System;
using System.Collections.Generic;

namespace GV.Domain.Entities
{
    public class SchedulesEntity
    {
        private readonly IList<DevicesEntity> _devices;
        private readonly IList<ScheduleServiceEntity> _scheduleServices;

        public SchedulesEntity()
        {
            _devices = new List<DevicesEntity>();
            _scheduleServices = new List<ScheduleServiceEntity>();
        }

        public virtual long ScheduleId { get; set; }
        public virtual long CustomerId { get; set; }
        public virtual string Name { get; set; }
        public virtual string Suffix { get; set; }
        public virtual DateTimeOffset? EffectiveDateTime { get; set; }
        public virtual DateTimeOffset? ExpiredDateTime { get; set; }
        public virtual int? Term { get; set; }
        public virtual decimal? ServiceAdjustment { get; set; }
        public virtual decimal MonthlyHwCost { get; set; }
        public virtual decimal MonthlySvcCost { get; set; }       
        public virtual decimal MonthlyContractCost { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }
        public virtual bool IsDeleted { get; set; }
        public virtual SchedulesEntity CoterminousSchedule { get; set; }
        public virtual DateTimeOffset?  ModifiedDateTime { get; set; }

        public virtual IEnumerable<DevicesEntity> Devices => _devices;
        public virtual IEnumerable<ScheduleServiceEntity> ScheduleServices => _scheduleServices;

        public virtual void AddDevice(DevicesEntity entity)
        {
            entity.Schedule = this;
            _devices.Add(entity);
        }

        public virtual void AddScheduleService(ScheduleServiceEntity service)
        {
            service.Schedule = this;
            _scheduleServices.Add(service);
        }
        
    }
}