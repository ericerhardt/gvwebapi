using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace GVWebapi.Models
{
    public class PeriodHistoryView
    {
      [Key]
      public long RevisionDataID { get; set; } // [RevisionDataID]
      public string ERPMeterGroupDesc { get; set; }//[ERPMeterGroupDesc]
      public Int64? ContractedVolume { get; set; }//[ContractedVolume]
      public Int64? ActualVolume { get; set; }//[ActualVolume]
      public decimal? CPPRate { get; set; }//[CPPRate]
      public Int64? VolumeOffset { get; set; }//[VolumeOffset]
      public DateTime? PeriodDate { get; set; }//[PeriodDate]
      public decimal? Credits { get; set; }//[Credits]
      public decimal? FBRContractBase { get; set; }//[FBRContractBase]
      public decimal? ClientContractBase { get; set; }//[ClientContractBase]
      public decimal? ClientCPP { get; set; }//[ClientCPP]
      public Int32? ERPContractID { get; set; }//[ERPContractID]
      public Int32? Rollover { get; set; }//[Rollover]
      public decimal? OverageExpense { get; set; }//[OverageExpense]
      public int? ERPMeterGroupID { get; set; }//[ERPMeterGroupID]
    
    }
}