//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace GVWebapi.RemoteData
{
    using System;
    using System.Collections.Generic;
    
    public partial class vw_RevisionInvoiceHistory
    {
        public int InvoiceID { get; set; }
        public int BillingMeterGroupID { get; set; }
        public string MeterGroup { get; set; }
        public Nullable<decimal> ContractVolume { get; set; }
        public Nullable<decimal> ActualVolume { get; set; }
        public string ID { get; set; }
        public Nullable<decimal> Overage { get; set; }
        public Nullable<decimal> CPP { get; set; }
        public Nullable<decimal> OverageCharge { get; set; }
        public int CreditAmount { get; set; }
        public decimal Rollover { get; set; }
        public Nullable<System.DateTime> OverageToDate { get; set; }
        public Nullable<System.DateTime> OverageFromDate { get; set; }
        public int CustomerID { get; set; }
        public Nullable<int> ContractMeterGroupID { get; set; }
        public int ContractID { get; set; }
        public string ContractMeterGroup { get; set; }
        public bool Void { get; set; }
        public decimal RolloverUsage { get; set; }
        public decimal RolloverUsageChange { get; set; }
    }
}
