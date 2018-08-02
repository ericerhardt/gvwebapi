using System;
using System.Collections.Generic;

namespace GV.Domain.Entities
{
    public class CyclesEntity
    {
        private readonly IList<CyclePeriodEntity> _cyclePeriods;
        private readonly IList<CycleReconciliationServicesEntity> _reconServices;

        public CyclesEntity(long customerId, DateTime startDate) : this()
        {
            CustomerId = customerId;
            StartDate = startDate;
            CreatedDateTime = DateTimeOffset.Now;
        }

        protected CyclesEntity()
        {
            _cyclePeriods = new List<CyclePeriodEntity>();
            _reconServices = new List<CycleReconciliationServicesEntity>();
        }

        public virtual long CycleId { get; set; }
        public virtual long CustomerId { get; set; }
        public virtual DateTime StartDate { get; set; }
        public virtual DateTime? EndDate { get; set; }
        public virtual bool InActive { get; set; }
        public virtual bool IsDeleted { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }
        public virtual DateTimeOffset? ModifiedDateTime { get; set; }
        public virtual bool InvisibleToClient { get; set; }
        public virtual bool IsReconciled { get; set; }

        public virtual IEnumerable<CyclePeriodEntity> CyclePeriods => _cyclePeriods;

        public virtual IEnumerable<CycleReconciliationServicesEntity> ReconServices => _reconServices;

        public virtual void AddNewPeriod(CyclePeriodEntity period)
        {
            period.Cycle = this;
            _cyclePeriods.Add(period);
        }

        public virtual void AddNewReconService(CycleReconciliationServicesEntity reconService)
        {
            reconService.Cycle = this;
            _reconServices.Add(reconService);
        }

    }
}