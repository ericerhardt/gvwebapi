using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models.Reconciliation
{
    public class DeviceSummaryModel
    {
        public int EquipmentID { get; set; }
        public string MeterGroup { get; set; }
        public decimal Cost { get; set; }
        public decimal Overage { get; set; }
    }
    public class DeviceCostsModel
    {
        
        public decimal Overage { get; set; }
        public decimal Cost { get; set; }
    }
}