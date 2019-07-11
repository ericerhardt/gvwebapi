using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models.CostAllocation
{
    public class AllocatedHardwareViewModel
    {
        public int CyclePeriodId { get; set; }
        public int ScheduleId { get; set; }
        public string ScheduleName { get; set; }
        public string CostCenter { get; set; }
        public string SerialNumber { get; set; }
        public string Model { get; set; }
        public string EquipmentNumber { get; set; }
        public string Status { get; set; }
        public string EquipmentType { get; set; }
        public string Location { get; set; }
        public string User { get; set; }
        public decimal? Hardware { get; set; }
        public decimal? HardwareTax { get; set; }
        public decimal? TotalCost { get; set; }
    }
}