using System;

namespace GV.Domain.Views
{
    public class ViewCsQuarterlyHistory
    {
        public int InvoiceId { get; set; }
        public DateTime Date { get; set; }
        public string InvoiceNumber { get; set; }
        public string ContractNumber { get; set; }
        public string CustomerNumber { get; set; }
        public string ContractMeterGroup { get; set; }
        public decimal GroupCopies { get; set; }
        public decimal CountedCopies { get; set; }
        public decimal CopyCredits { get; set; }
        public decimal CoveredCopies { get; set; }
        public decimal BillableCopies { get; set; }
        public decimal TotalChargeAmount { get; set; }
        public decimal EffectiveRate { get; set; }
        public string CreatorId { get; set; }
        public DateTime CreateDate { get; set; }
        public decimal RolloverUsage { get; set; }
        public decimal RolloverUsageChange { get; set; }
        public string CustomerName { get; set; }
        public string ContractMeterGroupDescription { get; set; }
        public int CustomerId { get; set; }
    }
}