using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using GVWebapi.Models.Schedules;
namespace GVWebapi.Models.CostAllocation
{
    public class AllocatedServicesViewModel
    {
        public long CyclePeriodId { get; set; }
        public long ScheduleId { get; set; }
        public string ScheduleName { get; set; }
        public decimal InstanceInvoiced { get; set; }
        public string CostCenter { get; set; }
        public List<MeterGroupCostCenter> MeterGroups { get; set; } = new List<MeterGroupCostCenter>();
        public decimal ServiceCost { get; set; }
        public decimal ServiceCostTax { get; set; }
        public decimal TotalCost { get; set; }
    }
}