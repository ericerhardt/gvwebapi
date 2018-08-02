using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class VisionModel
    {
        public long? RevisonDataID { get; set; }
        public int? ContractID {get;set;}
        public DateTime? PeriodDate { get; set; }
        public int? MeterGroupID { get; set; }
        public string MeterGroup { get; set; }
        public decimal? ContractVolume   { get; set; }
        public decimal? ActualVolume { get; set; }
        public decimal? Overage { get; set; }
        public decimal? RollOver { get; set; }
        public decimal?  CPP { get; set; }
        public decimal? CreditAmount { get; set; }
        public decimal? OverageCharge { get; set; }
        public DateTime? OverageToDate { get; set; }
        public int? InvoiceId { get; set; }
        public string Id { get; set; }
        public string month   { get; set; }
        public decimal? clientOverage { get; set; }
        public decimal? clientCPP { get; set; }
        public decimal? clientSavings { get; set; }
        public decimal? FPRBase { get; set; }
        public decimal? ClientBase { get; set; }
    }
}