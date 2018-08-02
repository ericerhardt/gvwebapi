using System.Collections.Generic;
using GV.Domain.Entities;

namespace GVWebapi.Models.Schedules
{
    public class EditScheduleTopModel
    {
        public static EditScheduleTopModel For(SchedulesEntity entity)
        {
            var model = new EditScheduleTopModel();
            if (entity == null) return model;
            model.Name = entity.Name;
            model.ScheduleId = entity.ScheduleId;
            model.MonthlySvcCost = entity.MonthlySvcCost;
            model.MonthlyHwCost = entity.MonthlyHwCost;
            return model;
        }

        private EditScheduleTopModel()
        {
        }

        public string Name { get; set; }
        public long ScheduleId { get; set; }
        public decimal MonthlyHwCost { get; set; }
        public decimal MonthlySvcCost { get; set; }
        public decimal TotalCost => MonthlySvcCost + MonthlyHwCost;
        public IList<SchedulesModel> ActiveSchedules {get; private set; }

        public void SetSchedule(IList<SchedulesModel> schedules)
        {
            ActiveSchedules = schedules;
        }
    }
}