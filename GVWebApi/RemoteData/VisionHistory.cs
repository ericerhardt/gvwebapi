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
    
    public partial class VisionHistory
    {
        public int VisionHistoryID { get; set; }
        public Nullable<int> ERPContractID { get; set; }
        public Nullable<System.DateTime> PeriodDate { get; set; }
        public string MeterTypeDesc { get; set; }
        public Nullable<int> QtrContractVolume { get; set; }
        public Nullable<int> QtrActualVolume { get; set; }
        public Nullable<double> OverageRate { get; set; }
        public Nullable<double> VolumeChange { get; set; }
        public Nullable<int> OverageVolume { get; set; }
        public Nullable<double> OverageExpense { get; set; }
        public Nullable<double> CostWOFPR { get; set; }
        public Nullable<double> NetOverageSavings { get; set; }
        public Nullable<double> PctOverageSavings { get; set; }
        public Nullable<int> VolumeOffset { get; set; }
    }
}