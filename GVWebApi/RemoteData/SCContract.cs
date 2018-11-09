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
    
    public partial class SCContract
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public SCContract()
        {
            this.SCContractMeterGroups = new HashSet<SCContractMeterGroup>();
        }
    
        public int ContractID { get; set; }
        public string ContractNumber { get; set; }
        public string ContractMajor { get; set; }
        public string ContractMinor { get; set; }
        public int CustomerID { get; set; }
        public Nullable<int> ContactID { get; set; }
        public string Contact { get; set; }
        public string Phone { get; set; }
        public string Fax { get; set; }
        public int BillToID { get; set; }
        public int ContractCodeID { get; set; }
        public int BillCodeID { get; set; }
        public Nullable<int> TaxCodeID { get; set; }
        public bool Taxable { get; set; }
        public Nullable<int> BillGroupID { get; set; }
        public System.DateTime StartDate { get; set; }
        public Nullable<System.DateTime> ExpDate { get; set; }
        public Nullable<int> ExpCopies { get; set; }
        public Nullable<decimal> AccumCopies { get; set; }
        public Nullable<System.DateTime> ExpCopyExpDate { get; set; }
        public Nullable<System.DateTime> BaseNextBillingDate { get; set; }
        public Nullable<System.DateTime> BaseLastBillingDate { get; set; }
        public Nullable<System.DateTime> BaseBilledThruDate { get; set; }
        public Nullable<int> BaseBillingCycleID { get; set; }
        public bool BaseArrears { get; set; }
        public Nullable<System.DateTime> BaseNextAccrualDate { get; set; }
        public Nullable<System.DateTime> BaseLastAccrualDate { get; set; }
        public Nullable<System.DateTime> BaseAccruedThruDate { get; set; }
        public Nullable<int> BaseAccrualCycleID { get; set; }
        public bool SumIndividualBaseRates { get; set; }
        public decimal BaseRate { get; set; }
        public int BaseRatePeriod { get; set; }
        public int BasePreBillingNoDays { get; set; }
        public int BasePreBillingNoMonths { get; set; }
        public Nullable<System.DateTime> OverageNextBillingDate { get; set; }
        public Nullable<System.DateTime> OverageLastBillingDate { get; set; }
        public Nullable<System.DateTime> OverageBilledThruDate { get; set; }
        public Nullable<int> OverageBillingCycleID { get; set; }
        public bool BSABillForServices { get; set; }
        public Nullable<double> BSALaborDiscount { get; set; }
        public Nullable<double> BSAMaterialsDiscount { get; set; }
        public decimal BSAMinimumBalance { get; set; }
        public decimal BSAMinimumBilling { get; set; }
        public Nullable<int> CoveredCopies { get; set; }
        public string LeaseSchedule { get; set; }
        public bool GroupInvoices { get; set; }
        public string PONumber { get; set; }
        public decimal MiscChargeAmount { get; set; }
        public int MiscChargeTaxFlag { get; set; }
        public string MiscChargeDescription { get; set; }
        public Nullable<int> MiscChargeGLID { get; set; }
        public Nullable<int> MiscChargeDeptID { get; set; }
        public bool MiscContinuous { get; set; }
        public string Remarks { get; set; }
        public string RemarksInternal { get; set; }
        public decimal UnearnedBalance { get; set; }
        public bool Bill { get; set; }
        public Nullable<System.DateTime> Activated { get; set; }
        public Nullable<System.DateTime> Terminated { get; set; }
        public Nullable<System.DateTime> Renewed { get; set; }
        public Nullable<int> RenewedID { get; set; }
        public int BranchID { get; set; }
        public Nullable<int> ContractLeaseID { get; set; }
        public int ExtendedTypeID { get; set; }
        public bool UseLeaseOnAllEquipment { get; set; }
        public Nullable<int> TermID { get; set; }
        public Nullable<int> SalesRepID { get; set; }
        public Nullable<int> JobID { get; set; }
        public Nullable<int> ChargeMethodID { get; set; }
        public Nullable<int> ChargeAccountID { get; set; }
        public bool SyncBillToAddress { get; set; }
        public string BillToAttn { get; set; }
        public string BillToAddress { get; set; }
        public string BillToCity { get; set; }
        public string BillToState { get; set; }
        public string BillToZip { get; set; }
        public string BillToCountry { get; set; }
        public bool UseAlternateOvgBillTo { get; set; }
        public Nullable<int> OvgBillToID { get; set; }
        public string OvgBillToAttn { get; set; }
        public string OvgBillToAddress { get; set; }
        public string OvgBillToCity { get; set; }
        public string OvgBillToState { get; set; }
        public string OvgBillToZip { get; set; }
        public string OvgBillToCountry { get; set; }
        public bool SyncOvgBillToAddress { get; set; }
        public Nullable<int> OvgBillToTermID { get; set; }
        public Nullable<int> OvgBillToChargeMethodID { get; set; }
        public Nullable<int> OvgBillToChargeAccountID { get; set; }
        public bool Renewable { get; set; }
        public Nullable<int> RenewalCycleID { get; set; }
        public Nullable<System.DateTime> BaseRateScheduleStartDate { get; set; }
        public int BaseRateScheduleStartDateType { get; set; }
        public Nullable<bool> BaseRateScheduleIsRenewalType { get; set; }
        public Nullable<int> NextBaseScheduleDetailID { get; set; }
        public Nullable<System.DateTime> NextBaseIncreaseDate { get; set; }
        public Nullable<System.DateTime> OvgRateScheduleStartDate { get; set; }
        public int OvgRateScheduleStartDateType { get; set; }
        public Nullable<bool> OvgRateScheduleIsRenewalType { get; set; }
        public Nullable<System.DateTime> NextOvgIncreaseDate { get; set; }
        public Nullable<int> NextOvgScheduleDetailID { get; set; }
        public Nullable<System.DateTime> BaseBillingStartDate { get; set; }
        public Nullable<System.DateTime> BaseAccrualStartDate { get; set; }
        public Nullable<System.DateTime> OverageBillingStartDate { get; set; }
        public bool EnableAgentUsageAlert { get; set; }
        public bool PooledUsageLimit { get; set; }
        public Nullable<int> ContractAdjCodeID { get; set; }
        public bool Active { get; set; }
        public int Locks { get; set; }
        public int DBFileHeaderID { get; set; }
        public string CreatorID { get; set; }
        public string UpdatorID { get; set; }
        public System.DateTime CreateDate { get; set; }
        public System.DateTime LastUpdate { get; set; }
        public byte[] timestamp { get; set; }
        public Nullable<int> ShTrackingConfigID { get; set; }
        public bool UseIndividualTaxCodes { get; set; }
        public Nullable<decimal> BSABillingDiscount { get; set; }
        public decimal BSABillingDiscountBalance { get; set; }
        public bool BSABillingDiscountMode { get; set; }
        public bool Accrual { get; set; }
        public bool AccrualAccountsLock { get; set; }
        public bool Billed { get; set; }
        public bool LockBillCode { get; set; }
        public bool UseBillCodeRevOnRenewal { get; set; }
        public int BSARebillType { get; set; }
        public Nullable<int> NoteID { get; set; }
        public int NoteFlag { get; set; }
        public Nullable<int> TerminationCodeID { get; set; }
        public int ContractStatusID { get; set; }
        public string OneTimeRemark { get; set; }
        public bool PrintOneTimeRemark { get; set; }
        public bool OvgBillToTaxable { get; set; }
        public Nullable<int> OvgBillToTaxCodeID { get; set; }
        public bool OvgBillToUseIndividualTaxCodes { get; set; }
        public int MiscChargeBillCycleFlag { get; set; }
        public bool AllowMeterEstimation { get; set; }
        public int ExpCopiesBase { get; set; }
        public int ExpAdjCopies { get; set; }
        public Nullable<System.DateTime> ExpCopiesActualExpDate { get; set; }
        public Nullable<System.DateTime> ExpDateInitial { get; set; }
        public bool ExpCopiesBillOverages { get; set; }
        public bool ExpCopiesRollOverages { get; set; }
        public int ExpCopiesMaxOveragesToRoll { get; set; }
        public int ExpCopiesActualOveragesToRoll { get; set; }
        public int CPPMinimumPages { get; set; }
        public decimal CPPRate { get; set; }
        public decimal CPPHardwareAmount { get; set; }
        public decimal CPPServiceRate { get; set; }
        public bool CombineLeaseAndBase { get; set; }
        public bool TimeBlockBased { get; set; }
        public Nullable<bool> MGBaseRateScheduleIsRenewalType { get; set; }
        public Nullable<System.DateTime> MGBaseRateScheduleStartDate { get; set; }
        public Nullable<int> MGBaseRateScheduleStartDateType { get; set; }
        public Nullable<System.DateTime> NextMGBaseIncreaseDate { get; set; }
        public Nullable<int> NextMGBaseScheduleDetailID { get; set; }
        public int ReportDefinitionGroupID { get; set; }
        public Nullable<int> ReportDefinitionContactID { get; set; }
        public Nullable<byte> ReportDefinitionDocSendMethodID { get; set; }
        public bool ReportDefinitionUseContactWithServiceInvoices { get; set; }
        public Nullable<int> OvgBillToReportDefinitionGroupID { get; set; }
        public Nullable<int> OvgBillToReportDefinitionContactID { get; set; }
        public Nullable<byte> OvgBillToReportDefinitionDocSendMethodID { get; set; }
        public bool OvgBillToReportDefinitionUseContactWithServiceInvoices { get; set; }
        public int NoteCount { get; set; }
        public bool IssueCredit { get; set; }
        public bool EnableBaseCycleConsolidation { get; set; }
        public bool EnableOverageCycleConsolidation { get; set; }
        public int DivisionID { get; set; }
        public int BalanceSheetGroupID { get; set; }
        public Nullable<int> LeaseDivisionID { get; set; }
        public Nullable<int> LeaseBalanceSheetGroupID { get; set; }
        public Nullable<System.DateTime> BaseBilledThruDateInitial { get; set; }
        public Nullable<System.DateTime> OverageBilledThruDateInitial { get; set; }
        public Nullable<System.DateTime> BaseAccruedThruDateInitial { get; set; }
        public int BaseRatePeriodType { get; set; }
        public decimal UnearnedInterestBalance { get; set; }
        public decimal UnearnedLeaseBalance { get; set; }
        public decimal UnearnedLeasePostTermBalance { get; set; }
        public string Message { get; set; }
        public bool MessReqAcknowledgment { get; set; }
        public Nullable<int> TaxExemptCodeID { get; set; }
        public Nullable<int> OvgBillToTaxExemptCodeID { get; set; }
        public Nullable<int> SLACodeID { get; set; }
        public int BaseDistributionCodeID { get; set; }
        public int MiscChargeTaxFlagID { get; set; }
    
        public virtual SCBillingCycle SCBillingCycle { get; set; }
        public virtual SCBillingCycle SCBillingCycle1 { get; set; }
        public virtual SCBillingCycle SCBillingCycle2 { get; set; }
        public virtual SCBillingCycle SCBillingCycle3 { get; set; }
        public virtual ARCustomer ARCustomer { get; set; }
        public virtual ARCustomer ARCustomer1 { get; set; }
        public virtual ARCustomer ARCustomer2 { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<SCContractMeterGroup> SCContractMeterGroups { get; set; }
    }
}
