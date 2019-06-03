using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models.Reports
{
    public class ContractedVolumeReport
    {
        public string Period { get; set; }
        public int MeterGroup { get; set; }
        public string MeterGroupDesc { get; set; }
        public decimal? ContractedVolume { get; set; }
        public decimal? ActualVolume { get; set; }
    }
}