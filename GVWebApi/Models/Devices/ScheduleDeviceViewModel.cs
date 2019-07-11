using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models.Devices
{
    public class ScheduleDeviceViewModel
    {
        public long ScheduleDeviceID { get; set; }
        public int EquipmentID { get; set; }
        public string EquipmentNumber { get; set; }
        public string SerialNumber { get; set; }
        public string Model { get; set; }
        public string ScheduleNumber { get; set; }
        public string Exhibit { get; set; }
        public string Location { get; set; }
        public int? LocationID { get; set; }
        public string User { get; set; }
        public string CostCenter { get; set; }
        public decimal MonthlyCost { get; set; }
        public bool Active { get; set; }

    }
}