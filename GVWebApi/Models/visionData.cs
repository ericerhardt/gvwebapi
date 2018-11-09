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
        public Decimal FPROverageCost { get; set; }
        public Decimal FPRCost { get; set; }
        public Decimal ClientOverageCost { get; set; }
        public Decimal ClientCost { get; set; }
        public Decimal Credits { get; set; }
        public Decimal Savings { get; set; }
        public Decimal Pct { get; set; }
       
    }
}
