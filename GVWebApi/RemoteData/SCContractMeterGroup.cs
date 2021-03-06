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
    
    public partial class SCContractMeterGroup
    {
        public int ContractMeterGroupID { get; set; }
        public string ContractMeterGroup { get; set; }
        public string Description { get; set; }
        public int ContractID { get; set; }
        public Nullable<decimal> CoveredCopies { get; set; }
        public Nullable<bool> ApplyToExpiration { get; set; }
        public Nullable<int> OverageTypeID { get; set; }
        public bool UseOverages { get; set; }
        public int MeterCount { get; set; }
        public Nullable<System.DateTime> NextOvgIncreaseDate { get; set; }
        public Nullable<int> NextOvgScheduleDetailID { get; set; }
        public Nullable<System.DateTime> OvgRateScheduleStartDate { get; set; }
        public Nullable<int> OvgRateScheduleStartDateType { get; set; }
        public Nullable<bool> OvgRateScheduleIsRenewalType { get; set; }
        public string CreatorID { get; set; }
        public string UpdatorID { get; set; }
        public System.DateTime CreateDate { get; set; }
        public System.DateTime LastUpdate { get; set; }
        public byte[] timestamp { get; set; }
        public int OverageMethodID { get; set; }
        public decimal CreditRate { get; set; }
        public decimal RolloverUsage { get; set; }
        public Nullable<int> BaseDistributionCodeID { get; set; }
        public decimal BilledCoveredCopiesBalance { get; set; }
        public bool BillMeterGroupBaseAmount { get; set; }
        public decimal BaseRatePerCopy { get; set; }
        public int RoundBaseAmountDigits { get; set; }
        public decimal UnearnedBalance { get; set; }
        public Nullable<bool> BaseRateScheduleIsRenewalType { get; set; }
        public Nullable<System.DateTime> BaseRateScheduleStartDate { get; set; }
        public Nullable<int> BaseRateScheduleStartDateType { get; set; }
        public Nullable<System.DateTime> NextBaseIncreaseDate { get; set; }
        public Nullable<int> NextBaseScheduleDetailID { get; set; }
        public Nullable<decimal> NextCoveredCopies { get; set; }
        public Nullable<System.DateTime> NextCoveredCopiesDate { get; set; }
        public int CoveredCopiesPer { get; set; }
        public decimal BilledCoveredCopiesBalanceAdjustment { get; set; }
        public Nullable<int> OverageBillingCycleID { get; set; }
        public Nullable<System.DateTime> OverageBilledThruDate { get; set; }
        public Nullable<System.DateTime> OverageNextBillingDate { get; set; }
        public Nullable<System.DateTime> OverageLastBillingDate { get; set; }
        public Nullable<System.DateTime> OverageBilledThruDateInitial { get; set; }
    
        public virtual SCBillingCycle SCBillingCycle { get; set; }
        public virtual SCContract SCContract { get; set; }
    }
}
