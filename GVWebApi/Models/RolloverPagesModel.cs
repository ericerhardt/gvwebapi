using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using GVWebapi.RemoteData;

namespace GVWebapi.Models
{
    public class RolloverPagesModel
    {
        public DateTime? Period { get; set; }
        public IEnumerable<Rollovers> Data { get; set; }
        public decimal TotalSavings { get; set; }
    }
    public class Rollovers
    {
        public string ERPMeterGroupDesc { get; set; }
        public Nullable<long> Rollover { get; set; }
        public decimal CPP { get; set; }
        public Nullable<decimal> Savings { get; set; }
    }
}