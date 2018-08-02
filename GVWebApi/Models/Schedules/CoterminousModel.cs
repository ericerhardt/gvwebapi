using System;
using GV.Domain.Entities;

namespace GVWebapi.Models.Schedules
{
    public class CoterminousModel
    {
        public static CoterminousModel For(long scheduleId, string name)
        {
            var model = new CoterminousModel();
            model.ScheduleId = scheduleId;
            model.Name = name;
            return model;
        }

        public static CoterminousModel For(SchedulesEntity schedule)
        {
            var model = new CoterminousModel();
            model.ScheduleId = schedule.ScheduleId;
            model.Name = schedule.Name;
            model.EffectiveDateTime = schedule.EffectiveDateTime;
            model.Term = schedule.Term;
            return model;
        }

        public long ScheduleId { get; set; }
        public string Name { get; set; }
        public bool IsSelected { get; set; }
        public DateTimeOffset? EffectiveDateTime {get; set; }
        public int? Term { get; set; }
    }
}