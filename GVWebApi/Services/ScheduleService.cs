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
        bool ScheduleExists(ScheduleSaveModel model);
        IList<SchedulesModel> GetAll(long customerId);
        void DeleteSchedule(long scheduleId);
        IList<CoterminousModel> GetCoterminous(long customerId);
        ScheduleDeleteModel CanDeleteSchedule(long scheduleId);
        ScheduleEditModel GetExistingSchedule(long scheduleId);
        void UpdateSchedule(ScheduleSaveModel model);
        void UpdateMonthyHardwareCost(decimal value, long scheduleId);
        IList<SchedulesModel> GetActiveSchedules(long deviceId);
        IList<SchedulesModel> GetAcitveSchedulesByScheduleId(long scheduleId);
        IList<SchedulesModel> GetAcitveSchedulesByCustomer(long CustomerId);
        SchedulesModel GetSchedule(long scheduleId);
      
    }

    public class ScheduleService : IScheduleService
    {
        private readonly IScheduleServicesService _scheduleServicesService;
        private readonly IDeviceService _deviceService;
        private readonly IRepository _repository;
        private readonly ICoFreedomDeviceService _coFreedomDeviceService;

        public ScheduleService(IRepository repository, ICoFreedomDeviceService coFreedomDeviceService, IScheduleServicesService scheduleServicesService, IDeviceService deviceService)
        {
            _repository = repository;
            _coFreedomDeviceService = coFreedomDeviceService;
            _scheduleServicesService = scheduleServicesService;
            _deviceService = deviceService;
        }
        public bool ScheduleExists(ScheduleSaveModel model)
        {
            var hasschedule = _repository.Find<SchedulesEntity>().Where(x => x.CustomerId == model.CustomerId && x.Name == model.Name && x.IsDeleted == false).ToList();
            if (hasschedule.Count() > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public void AddSchedule(ScheduleSaveModel model)
        {

            var schedulesEntity = new SchedulesEntity();
            schedulesEntity.CustomerId = model.CustomerId;
            schedulesEntity.Name = model.Name;
            schedulesEntity.Suffix = model.Suffix;
            schedulesEntity.Term = model.Term;
            schedulesEntity.EffectiveDateTime = model.EffectiveDateTime;
            schedulesEntity.ExpiredDateTime = GetExpiredDateTime(model.EffectiveDateTime, model.Term);
            schedulesEntity.MonthlyContractCost = model.MonthlyContractCost;
            schedulesEntity.CreatedDateTime = DateTimeOffset.Now;

            if (model.CoterminousId != 0)
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

            var results = _repository.Add(schedulesEntity);
            if (model.CoterminousId != 0)
            {
                var meterGroups = _scheduleServicesService.GetMeterGroups(model.CoterminousId.Value);
                foreach (var metergroup in meterGroups)
                {
                   var scheduleServiceModel = new ScheduleServiceEntity();
                    scheduleServiceModel.MeterGroup = metergroup.MeterGroup;
                    scheduleServiceModel.BaseCpp = metergroup.BaseCpp;
                    scheduleServiceModel.OverageCpp = metergroup.OverageCpp;
                    scheduleServiceModel.Schedule = results;
                    _repository.Add(scheduleServiceModel);
                }
            }
             var TotalCost = _deviceService.DeviceTotalCost(results.ScheduleId);
             UpdateMonthyHardwareCost(TotalCost, results.ScheduleId);
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
 
                foreach (var schedulesModel in schedules)
                {
                    
                    schedulesModel.SetDeviceCount(_coFreedomDeviceService.GetScheduleDeviceCount(schedulesModel.Name));
                }
            }
            return schedules.OrderBy(x => x.Suffix).ToList();
        }

        public void DeleteSchedule(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);
            schedule.IsDeleted = true;
            var TotalCost = _deviceService.DeviceTotalCost(schedule.ScheduleId);
            UpdateMonthyHardwareCost(TotalCost, schedule.ScheduleId);
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
        public void UpdateMonthyHardwareCost(decimal value,long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);
            if (schedule == null) return;
            schedule.MonthlyHwCost = value;
            _repository.Add(schedule);
        }
        private string GetSufix(string ScheduleName)
        {
            var results = String.Empty;
            if(ScheduleName.Length > 13)
            {
                results = ScheduleName.Substring(ScheduleName.Length - 6, 3);
            } else
            {
                results = ScheduleName.Substring(ScheduleName.Length - 3, 3);
            }
         
            return results;
        }
        public void UpdateSchedule(ScheduleSaveModel model)
        {
            var schedule = _repository.Get<SchedulesEntity>(model.ScheduleId);
            if (schedule == null) return;
            schedule.Suffix = Int32.Parse(GetSufix(model.Name));
            schedule.Name = model.Name;
            schedule.MonthlyContractCost = model.MonthlyContractCost;
            schedule.ModifiedDateTime = DateTimeOffset.Now;

            if (model.CoterminousId.HasValue)
            {
              
                schedule.EffectiveDateTime = model.EffectiveDateTime;
                schedule.Term = model.Term;
                schedule.CoterminousSchedule = _repository.Load<SchedulesEntity>(model.CoterminousId.Value);
                schedule.ExpiredDateTime = schedule.CoterminousSchedule.ExpiredDateTime;
            }
            else
            {
                schedule.EffectiveDateTime = model.EffectiveDateTime;
                schedule.Term = model.Term;
                schedule.ExpiredDateTime = GetExpiredDateTime(model.EffectiveDateTime, model.Term);
            }
            var TotalCost = _deviceService.DeviceTotalCost(model.ScheduleId);
            UpdateMonthyHardwareCost(TotalCost, model.ScheduleId);
        }

        public IList<SchedulesModel> GetActiveSchedules(long deviceId)
        {
            var device = _coFreedomDeviceService.GetCoFreedomDevice(deviceId);

            return _repository.Find<SchedulesEntity>()
                .Where(x => x.IsDeleted == false)
                .Where(x => x.CustomerId == device.CustomerID)
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
        public IList<SchedulesModel> GetAcitveSchedulesByCustomer(long CustomerId)
        {
           
            return _repository.Find<SchedulesEntity>()
                .Where(x => x.IsDeleted == false)
                .Where(x => x.CustomerId ==  CustomerId)
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