using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class CostAllocationSettingsViewModel
    {
        public int SettingsMeterGroupID { get; set; }
        public int CustomerID { get; set; }
        public int MeterGroupID { get; set; }
        public string MeterGroupDesc { get; set; }
        public decimal ExcessCPP { get; set; }
    }
}