using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class RolloverUsageModel
    {
        public string ContractMeterGroup { get; set; }
        public int ContractMeterGroupID { get; set; }
        public int ContractID { get; set; }
        public decimal RolloverUsage { get; set; }
    }
}