using System;

namespace GVWebapi.Models.Schedules
{
    public class ScheduleSaveModel
    {
        public string Name { get; set; }
        public int Suffix { get; set; }
        public DateTimeOffset? EffectiveDateTime {get; set; }
        public DateTimeOffset? ExpiredDateTime {get; set; }  
        public decimal MonthlyContractCost { get; set; }
        public int? Term { get; set; }
        public decimal? ServiceAdjustment { get; set; }
        public int CustomerId { get; set; }
        public long? CoterminousId { get; set; }
        public long ScheduleId { get; set; }
    }
}