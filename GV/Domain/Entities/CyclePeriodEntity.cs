using System;
using System.Collections.Generic;

namespace GV.Domain.Entities
{
    public class CyclePeriodEntity
    {
        private readonly IList<CyclePeriodSchedulesEntity> _periodSchedules;

        public CyclePeriodEntity(CyclesEntity cycle) : this()
        {
            Cycle = cycle;
            CreatedDateTime = DateTimeOffset.Now;
        }

        protected CyclePeriodEntity()
        {
            _periodSchedules = new List<CyclePeriodSchedulesEntity>();
        }

        public virtual long CyclePeriodId { get; set; }
        public virtual CyclesEntity Cycle { get; set; }
        public virtual DateTime Period { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }
        public virtual DateTimeOffset? ModifiedDateTime { get; set; }
        public virtual bool IsDeleted { get; set; }
        public virtual string InvoiceNumber { get; set; }

        public virtual IEnumerable<CyclePeriodSchedulesEntity> PeriodSchedules => _periodSchedules;

        public virtual void AddPeriodSchedule(CyclePeriodSchedulesEntity cyclePeriodSchedule)
        {
            cyclePeriodSchedule.CyclePeriod = this;
            _periodSchedules.Add(cyclePeriodSchedule);
        }
    }
}