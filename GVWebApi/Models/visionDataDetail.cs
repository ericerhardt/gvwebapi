using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class VisionDataDetail
    {
        public int          ERPContractID{get;set;}
        public DateTime     PeriodDate { get; set; }
        public String       ClientPeriodDates { get; set; }
        public String       MeterGroup { get; set; }
        public Double       ContractVolume { get; set; }
        public Double       AdjustedVolume { get; set; }
        public Double       ActualVolume { get; set; }
        public Double       VolumeOffset { get; set; }
        public Int32        OverageVolume { get; set; }
        public Double       NetChange { get; set; }
        public Double       OverageCost { get; set; }
        public Double       RolloverVolume { get; set; }
        public Decimal       CPP { get; set; } 
        public Double       ClientCPP { get; set; } 
        public Double       CreditAmount { get; set; }
        public String       month { get; set; }
        public Double       clientOverage { get; set; }
        public Double       Credits { get; set; }
        public Double       Savings { get; set; }
        public Double       PctSavings { get; set; }                            
  
    }
}
