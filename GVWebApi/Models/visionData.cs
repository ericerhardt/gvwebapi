using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class VisionData
    {
        public int ERPContractID{get;set;}
        public String ClientPeriodDates { get; set; }
        public DateTime ClientStartDate { get; set; }
        public DateTime ClientPeriodDate  { get; set; }
        public Double FPROverageCost { get; set; }
        public Double PreFPROverageCost { get; set; }
        public Double FPRCost { get; set; }
        public Double ClientOverageCost { get; set; }
        public Double ClientCost { get; set; }
        public Double Credits { get; set; }
        public Double Savings { get; set; }
        public Double Pct { get; set; }
        public Double? OverageCharge { get; set; }   
    }
}
