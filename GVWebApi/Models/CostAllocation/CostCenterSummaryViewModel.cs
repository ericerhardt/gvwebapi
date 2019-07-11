using GVWebapi.Models.Schedules;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models.CostAllocation
{
    public class CostCenterSummaryViewModel
    {
        public int RowNumber { get; set; }
        public string CostCenter { get; set; }
        public decimal Hardware { get; set; }
        public decimal HardwareTax { get; set; }
        public decimal Service { get; set; }
        public decimal ServiceTax { get; set; }
        public decimal CalculatedTax => Service * (ServiceTax / 100);
        public decimal InstanceInvoiced { get; set; }
        public List<MeterGroupCostCenter> MeterGroups { get; set; } = new List<MeterGroupCostCenter>();
        public decimal Adjustments { get; set; }
        public decimal SubTotal => (Hardware + HardwareTax + Service + CalculatedTax);
        public decimal TotalCost => (Hardware + HardwareTax + Service + CalculatedTax) * InstanceInvoiced;
    }
    public class ReconcileCostCenterSummary {

        public IList<CostCenterSummaryViewModel> summaries { get; set; } = new List<CostCenterSummaryViewModel>();
        public IList<MeterGroup> MeterGroups { get; set; } = new List<MeterGroup>();
    }

}