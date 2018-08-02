namespace GV.Domain.Entities
{
    public class CyclePeriodSchedulesEntity
    {
        public virtual long CyclePeriodScheduleId { get; set; }
        public virtual SchedulesEntity Schedule { get; set; }
        public virtual  CyclePeriodEntity CyclePeriod { get;set; }
        public virtual decimal InstancesInvoiced { get; set; }
    }
}