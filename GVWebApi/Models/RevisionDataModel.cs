using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public partial class RevisionDataModel
    {
        public int RevisionID { get; set; }
        public int InvoiceID { get; set; }
        public string ContractMeterGroup { get; set; }
        public string MeterGroup { get; set; }
        public Nullable<decimal> ContractVolume { get; set; }
        public Nullable<decimal> ActualVolume { get; set; }
        public Nullable<decimal> Overage { get; set; }
        public Nullable<decimal> CPP { get; set; }
        public Nullable<decimal> OverageCharge { get; set; }
        public decimal? CreditAmount { get; set; }
        public int? Rollover { get; set; }
        public Nullable<System.DateTime> OverageToDate { get; set; }
        public Nullable<System.DateTime> OverageFromDate { get; set; }
        public int CustomerID { get; set; }
        public Nullable<int> ContractMeterGroupID { get; set; }
        public int ContractID { get; set; }
       
       
    }
}