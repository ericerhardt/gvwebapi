using System;

namespace GVWebapi.Models.Reconciliation
{
    public class CycleSummaryModel
    {
        public long CyclePeriodId { get; set; }
        public DateTime Period { get; set; }
        public decimal AllocatedHardware { get; set; }
        public decimal UnallocatedService { get; set; }
        public decimal SubTotal => AllocatedHardware + UnallocatedService;
    }
}