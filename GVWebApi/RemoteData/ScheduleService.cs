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
    
    public partial class ScheduleService
    {
        public long ScheduleServiceId { get; set; }
        public long ScheduleId { get; set; }
        public string MeterGroup { get; set; }
        public int ContractedPages { get; set; }
        public decimal BaseCpp { get; set; }
        public decimal OverageCpp { get; set; }
        public decimal Cost { get; set; }
        public bool RemovedFromSchedule { get; set; }
        public bool IsDeleted { get; set; }
        public System.DateTimeOffset CreatedDateTime { get; set; }
        public Nullable<System.DateTimeOffset> ModifiedDateTime { get; set; }
        public Nullable<int> ContractMeterGroupID { get; set; }
    
        public virtual Schedule Schedule { get; set; }
    }
}
