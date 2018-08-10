using System;
using System.Collections.Generic;
using System.Linq;
using GV.Domain;
using GV.Domain.Entities;
using GVWebapi.Models.Schedules;

namespace GVWebapi.Services
{
    public interface IScheduleService
    {
        void AddSchedule(ScheduleSaveModel model);
        IList<SchedulesModel> GetAll(long customerId);
        void DeleteSchedule(long scheduleId);
        IList<CoterminousModel> GetCoterminous(long customerId);
        ScheduleDeleteModel CanDeleteSchedule(long scheduleId);
        ScheduleEditModel GetExistingSchedule(long scheduleId);
        void UpdateSchedule(ScheduleSaveModel model);
        IList<SchedulesModel> GetActiveSchedules(long deviceId);
        IList<SchedulesModel> GetAcitveSchedulesByScheduleId(long scheduleId);
        SchedulesModel GetSchedule(long scheduleId);
    }

    public class ScheduleService : IScheduleService
    {
        private readonly IRepository _repository;
        private readonly ICoFreedomDeviceService _coFreedomDeviceService;

        public ScheduleService(IRepository repository, ICoFreedomDeviceService coFreedomDeviceService)
        {
            _repository = repository;
            _coFreedomDeviceService = coFreedomDeviceService;
        }

        public void AddSchedule(ScheduleSaveModel model)
        {
            var schedulesEntity = new SchedulesEntity();
            schedulesEntity.CustomerId = model.CustomerId;
            schedulesEntity.Name = model.Name;
            schedulesEntity.Term = model.Term;
            schedulesEntity.EffectiveDateTime = model.EffectiveDateTime;
            schedulesEntity.ExpiredDateTime = GetExpiredDateTime(model.EffectiveDateTime, model.Term);
            schedulesEntity.MonthlySvcCost = model.MonthlySvcCost;
            schedulesEntity.MonthlyHwCost = model.MonthlyHwCost;
            schedulesEntity.CreatedDateTime = DateTimeOffset.Now;

            if (model.CoterminousId.HasValue)
            {
                schedulesEntity.CoterminousSchedule = _repository.Load<SchedulesEntity>(model.CoterminousId.Value);
            }
            else
            {
                schedulesEntity.EffectiveDateTime = model.EffectiveDateTime;
                schedulesEntity.Term = model.Term;
                schedulesEntity.ExpiredDateTime = model.EffectiveDateTime?.AddMonths(model.Term ?? 0);
                schedulesEntity.ExpiredDateTime = GetExpiredDateTime(model.EffectiveDateTime, model.Term);
            }

            _repository.Add(schedulesEntity);
        }

        private static DateTimeOffset? GetExpiredDateTime(DateTimeOffset? effectiveDate, int? term)
        {
            if (effectiveDate.HasValue == false || term.HasValue == false) return null;
            var totalDays = term.Value * 30.42;
            return effectiveDate.Value.AddDays(totalDays);
        }

        public IList<SchedulesModel> GetAll(long customerId)
        {
            var schedules = _repository.Find<SchedulesEntity>()
                .Where(x => x.IsDeleted == false)
                .Where(x => x.CustomerId == customerId)
                .Select(x => SchedulesModel.For(x))
                .ToList();

            if (schedules.Count > 0)
            {

            
                var schedulesAndDevices = _coFreedomDeviceService.GetDeviceCount(customerId);

                foreach (var schedulesModel in schedules)
                {
                   if(schedulesAndDevices.ContainsKey(schedulesModel.Name) == false) continue;
                    schedulesModel.SetDeviceCount(schedulesAndDevices[schedulesModel.Name]);
                }
            }
            return schedules;
        }

        public void DeleteSchedule(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);
            schedule.IsDeleted = true;
        }

        public IList<CoterminousModel> GetCoterminous(long customerId)
        {
            var databaseList = _repository.Find<SchedulesEntity>()
                .Where(x => x.IsDeleted == false)
                .Where(x => x.ExpiredDateTime >= DateTimeOffset.Now)
                .Where(x => x.CustomerId == customerId)
                .OrderBy(x => x.Name)
                .Select(x => CoterminousModel.For(x))
                .ToList();

            var models = new List<CoterminousModel>();
            models.Insert(0, CoterminousModel.For(0, "NO"));

            var itemsThatStartWithN = databaseList
                .Where(x => x.Name.StartsWith("n", StringComparison.OrdinalIgnoreCase))
                .OrderBy(x => x.Name);

            models.AddRange(itemsThatStartWithN);

            var itemsThatDidNotStartWithN = databaseList
                .Where(x => x.Name.StartsWith("n", StringComparison.OrdinalIgnoreCase) == false)
                .OrderBy(x => x.Name);

            models.AddRange(itemsThatDidNotStartWithN);

            SetDefaultSelected(models);

            return models;
        }

        public ScheduleDeleteModel CanDeleteSchedule(long scheduleId)
        {
            var schedules = _repository.Find<SchedulesEntity>()
                .Where(x => x.CoterminousSchedule.ScheduleId == scheduleId)
                .Where(x => x.IsDeleted == false)
                .ToList();

            var deleteModel = new ScheduleDeleteModel();

            if (schedules.Count == 0)
            {
                deleteModel.CanDelete = true;
            }
            else
            {
                deleteModel.ScheduleName = string.Join(",", schedules.Select(x => x.Name));
            }

            return deleteModel;
        }

        public ScheduleEditModel GetExistingSchedule(long scheduleId)
        {
            var model = new ScheduleEditModel();

            var schedulesEnity = _repository.Get<SchedulesEntity>(scheduleId);
            model.Schedule = SchedulesModel.For(schedulesEnity);
            model.Schedule.CoterminousScheduleId = schedulesEnity.CoterminousSchedule?.ScheduleId;
            model.CoterminousModels = GetCoterminous(schedulesEnity.CustomerId);

            return model;
        }

        public void UpdateSchedule(ScheduleSaveModel model)
        {
            var schedule = _repository.Get<SchedulesEntity>(model.ScheduleId);
            if (schedule == null) return;

            schedule.Name = model.Name;
            schedule.MonthlySvcCost = model.MonthlySvcCost;
            schedule.MonthlyHwCost = model.MonthlyHwCost;
            schedule.ModifiedDateTime = DateTimeOffset.Now;

            if (model.CoterminousId.HasValue)
            {
                schedule.ExpiredDateTime = null;
                schedule.EffectiveDateTime = null;
                schedule.Term = null;
                schedule.CoterminousSchedule = _repository.Load<SchedulesEntity>(model.CoterminousId.Value);
            }
            else
            {
                schedule.EffectiveDateTime = model.EffectiveDateTime;
                schedule.Term = model.Term;
                schedule.ExpiredDateTime = GetExpiredDateTime(model.EffectiveDateTime, model.Term);
            }
        }

        public IList<SchedulesModel> GetActiveSchedules(long deviceId)
        {
            var device =_repository.Get<DevicesEntity>(deviceId);

            return _repository.Find<SchedulesEntity>()
                .Where(x => x.IsDeleted == false)
                .Where(x => x.CustomerId == device.CustomerId)
                .Select(x => SchedulesModel.For(x))
                .ToList();
        }

        public IList<SchedulesModel> GetAcitveSchedulesByScheduleId(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);
            return _repository.Find<SchedulesEntity>()
                .Where(x => x.IsDeleted == false)
                .Where(x => x.CustomerId == schedule.CustomerId)
                .Select(x => SchedulesModel.For(x))
                .ToList();
                
        }

        public SchedulesModel GetSchedule(long scheduleId)
        {
            return SchedulesModel.For(_repository.Get<SchedulesEntity>(scheduleId));
        }

        private static void SetDefaultSelected(List<CoterminousModel> models)
        {
            if (models.Count == 1)
                models.First().IsSelected = true;

            var itemStartingWithN = models.Where(x => x.Name.StartsWith("n", StringComparison.OrdinalIgnoreCase)).ToList();

            if (itemStartingWithN.Count == 1)
                itemStartingWithN.First().IsSelected = true;

            var firstItemAfterNo = itemStartingWithN.SkipWhile(x => x.Name.ToLower() == "no").FirstOrDefault();
            if (firstItemAfterNo == null) return;
            firstItemAfterNo.IsSelected = true;
        }
    }
}