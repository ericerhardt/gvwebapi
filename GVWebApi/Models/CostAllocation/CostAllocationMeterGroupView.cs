using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models.CostAllocation
{
    public class CostAllocationMeterGroupView
    {
        public int SettingsMeterGroupID { get; set; }
        public int CustomerID { get; set; }
        public string MeterGroupName { get; set; }
        public int ContractMeterGroupID { get; set; }
        public long MeterGroupID { get; set; }
        public decimal ExcessCPP { get; set; }
    }
    
}