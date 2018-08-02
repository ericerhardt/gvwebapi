namespace GV.Domain.Entities
{
    public class CycleReconciliationServicesEntity
    {
        public CycleReconciliationServicesEntity(string meterGroup)
        {
            MeterGroup = meterGroup;
        }

        protected CycleReconciliationServicesEntity()
        {
        }

        public virtual long CycleReconciliationServiceId { get; set; }
        public virtual CyclesEntity Cycle { get; set; }
        public virtual string MeterGroup { get; set; }
        public virtual decimal Credit { get; set; }
    }
}