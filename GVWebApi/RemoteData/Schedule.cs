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
    
    public partial class Schedule
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public Schedule()
        {
            this.CyclePeriodSchedules = new HashSet<CyclePeriodSchedule>();
            this.Devices = new HashSet<Device>();
            this.ScheduleServices = new HashSet<ScheduleService>();
            this.Schedules1 = new HashSet<Schedule>();
        }
    
        public long ScheduleId { get; set; }
        public long CustomerId { get; set; }
        public string Name { get; set; }
        public Nullable<int> Suffix { get; set; }
        public Nullable<System.DateTimeOffset> EffectiveDateTime { get; set; }
        public Nullable<System.DateTimeOffset> ExpiredDateTime { get; set; }
        public Nullable<int> Term { get; set; }
        public Nullable<decimal> MonthlyHwCost { get; set; }
        public Nullable<decimal> MonthlySvcCost { get; set; }
        public Nullable<decimal> ServiceAdjustment { get; set; }
        public Nullable<decimal> MonthlyContractCost { get; set; }
        public System.DateTimeOffset CreatedDateTime { get; set; }
        public bool IsDeleted { get; set; }
        public Nullable<long> CoterminousScheduleId { get; set; }
        public Nullable<System.DateTimeOffset> ModifiedDateTime { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CyclePeriodSchedule> CyclePeriodSchedules { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Device> Devices { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<ScheduleService> ScheduleServices { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Schedule> Schedules1 { get; set; }
        public virtual Schedule Schedule1 { get; set; }
    }
}
