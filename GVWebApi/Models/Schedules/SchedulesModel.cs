using System;
using System.Data;
using GV.Domain.Entities;
using GVWebapi.Helpers;

namespace GVWebapi.Models.Schedules
{
    public class SchedulesModel
    {
        public static SchedulesModel For(SchedulesEntity scheduleEntity)
        {
            var schedule = new SchedulesModel();
            schedule.ScheduleId = scheduleEntity.ScheduleId;
            schedule.EffectiveDateTime = scheduleEntity.EffectiveDateTime;
            schedule.ExpiredDateTime = scheduleEntity.ExpiredDateTime;
            schedule.Term = scheduleEntity.Term;
            schedule.MonthlyHwCost = scheduleEntity.MonthlyHwCost;
            schedule.MonthlySvcCost = scheduleEntity.MonthlySvcCost;
            schedule.MonthlyContractCost = scheduleEntity.MonthlyContractCost;
            schedule.CreatedDateTime = scheduleEntity.CreatedDateTime;
            schedule.Name = scheduleEntity.Name;
            if (scheduleEntity.CoterminousSchedule != null)
                SetCoterminousParent(schedule, scheduleEntity);
            return schedule;
        }

        public string Name { get; set; }
        public long ScheduleId { get; set; }
        public DateTimeOffset? EffectiveDateTime { get; set; }
        public DateTimeOffset? ExpiredDateTime { get; set; }
        public int? Term { get; set; }
        public decimal? ServiceAdjustment { get; set; }
        public decimal MonthlyHwCost { get; set; }
        public decimal MonthlySvcCost { get; set; }
        public decimal MonthlyContractCost { get; set; }
        public DateTimeOffset CreatedDateTime { get; set; }
        public long? CoterminousScheduleId { get; set; }
        //this is used in the client
        public decimal TotalCost => MonthlySvcCost + MonthlyHwCost;
        public int DeviceCount { get; private set;}
        public void SetDeviceCount(int deviceCount) => DeviceCount = deviceCount;

        private static void SetCoterminousParent(SchedulesModel model, SchedulesEntity entity)
        {
            if (entity.CoterminousSchedule != null)
            {
                model.EffectiveDateTime = entity.CoterminousSchedule.EffectiveDateTime;
                model.ExpiredDateTime = entity.CoterminousSchedule.ExpiredDateTime;
                model.Term = entity.CoterminousSchedule.Term;
            }
            else
            {
                SetCoterminousParent(model, entity.CoterminousSchedule);
            }
        }
    }
}