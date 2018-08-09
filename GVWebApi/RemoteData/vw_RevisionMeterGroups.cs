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
    
    public partial class vw_RevisionMeterGroups
    {
        public int InvoiceID { get; set; }
        public int ContractID { get; set; }
        public System.DateTime Date { get; set; }
        public string InvoiceNumber { get; set; }
        public string ContractNumber { get; set; }
        public string CustomerNumber { get; set; }
        public string ContractMeterGroup { get; set; }
        public Nullable<decimal> GroupCopies { get; set; }
        public Nullable<decimal> CountedCopies { get; set; }
        public Nullable<decimal> CoveredCopies { get; set; }
        public Nullable<decimal> BillableCopies { get; set; }
        public Nullable<decimal> TotalChargeAmount { get; set; }
        public Nullable<decimal> EffectiveRate { get; set; }
        public string CreatorID { get; set; }
        public Nullable<System.DateTime> CreateDate { get; set; }
        public decimal RolloverUsage { get; set; }
        public decimal RolloverUsageChange { get; set; }
        public string CustomerName { get; set; }
        public string ContractMeterGroupDescription { get; set; }
        public Nullable<int> ContractMeterGroupID { get; set; }
        public Nullable<System.DateTime> OverageToDate { get; set; }
        public int BillingMeterGroupID { get; set; }
    }
}