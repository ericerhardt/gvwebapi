using System;

namespace GVWebapi.Models.Reconciliation
{
    public class CycleSummaryModel
    {
        public long CyclePeriodId { get; set; }
        public DateTime Period { get; set; }
        public decimal AllocatedHardware { get; set; }
        public decimal AllocatedHardwareTax { get; set; }
        public decimal CalculatedHardwarePlusTax => AllocatedHardware + AllocatedHardwareTax;
        public decimal AllocatedService { get; set; }
        public decimal AllocatedServiceTax { get; set; }
        public decimal CalculatedServiceTax => AllocatedService * (AllocatedServiceTax / 100);
        public decimal CalculatedServicePlusTax => AllocatedService + CalculatedServiceTax;
        public decimal Adjustments { get; set; }
        public decimal UnallocatedService { get; set; }
        public decimal UnallocatedServiceTax { get; set; }
        public decimal Total => CalculatedHardwarePlusTax + CalculatedServicePlusTax + Adjustments + UnallocatedService + UnallocatedServiceTax;
    }
}