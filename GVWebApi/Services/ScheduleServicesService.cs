using System;
using System.Collections.Generic;
using System.Linq;
using GV.CoFreedomDomain;
using GV.CoFreedomDomain.Entities;
using GV.Domain;
using GV.Domain.Entities;
using GVWebapi.Controllers;

namespace GVWebapi.Services
{
    public interface IScheduleServicesService
    {
        List<ScheduleServiceModel> GetMeterGroups(long scheduleId);
        void SaveSchedules(IList<ScheduleServiceSaveModel> scheduleServices);
        void UpdateRemovedFromSchedule(ScheduleServiceCheckChangedModel model);
    }

    public class ScheduleServicesService : IScheduleServicesService
    {
        private readonly ICoFreedomRepository _coFreedomRepository;
        private readonly IRepository _repository;
        private readonly IUnitOfWork _unitOfWork;

        public ScheduleServicesService(IRepository repository, ICoFreedomRepository coFreedomRepository, IUnitOfWork unitOfWork)
        {
            _repository = repository;
            _coFreedomRepository = coFreedomRepository;
            _unitOfWork = unitOfWork;
        }

        public List<ScheduleServiceModel> GetMeterGroups(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);

            var coFreedomMeterGroups = GetCoFreedomMeterGroups(schedule.CustomerId);

            foreach (var meterGroup in coFreedomMeterGroups)
            {
                var existingService = schedule.ScheduleServices.FirstOrDefault(x => x.MeterGroup == meterGroup);
                if (existingService == null)
                {
                    //doesn't exist in GV stuff
                    var scheduleService = new ScheduleServiceEntity(meterGroup);
                    schedule.AddScheduleService(scheduleService);
                }
            }

            _unitOfWork.Commit();

            return schedule.ScheduleServices
                .Where(x => x.IsDeleted == false)
                .Select(x => new ScheduleServiceModel
                {
                    ScheduleServiceId = x.ScheduleServiceId,
                    MeterGroup = x.MeterGroup,
                    ContractedPages = x.ContractedPages,
                    BaseCpp = x.BaseCpp,
                    OverageCpp = x.OverageCpp,
                    Cost = x.Cost,
                    RemovedFromSchedule = x.RemovedFromSchedule
                })
                .ToList();
        }

        public void SaveSchedules(IList<ScheduleServiceSaveModel> scheduleServices)
        {
            foreach (var scheduleServiceSaveModel in scheduleServices)
            {
                var scheduleService = _repository.Get<ScheduleServiceEntity>(scheduleServiceSaveModel.ScheduleServiceId);
                scheduleService.ContractedPages = scheduleServiceSaveModel.ContractedPages;
                scheduleService.BaseCpp = scheduleServiceSaveModel.BaseCpp;
                scheduleService.OverageCpp = scheduleServiceSaveModel.OverageCpp;
                scheduleService.ModifiedDateTime = DateTimeOffset.Now;
                scheduleService.SetCost();    
            }
        }

        public void UpdateRemovedFromSchedule(ScheduleServiceCheckChangedModel model)
        {
            var scheduleService = _repository.Get<ScheduleServiceEntity>(model.ScheduleServiceId);
            scheduleService.RemovedFromSchedule = !model.RemovedFromSchedule;
            scheduleService.ModifiedDateTime = DateTimeOffset.Now;
        }

        private IEnumerable<string> GetCoFreedomMeterGroups(long customerId)
        {
            return _coFreedomRepository.Find<ScEquipmentEntity>()
                .Where(x => x.CustomerId == customerId)
                .SelectMany(x => x.ContractDetails)
                .SelectMany(x => x.Contract.ContractMeterGroups)
                .Select(x => x.ContractMeterGroup)
                .Distinct()
                .ToList();
        }
    }

    public class ScheduleServiceModel
    {
        public long ScheduleServiceId { get; set; }
        public string MeterGroup { get; set; }
        public int ContractedPages { get; set; }
        public decimal BaseCpp { get; set; }
        public decimal OverageCpp { get; set; }
        public decimal Cost { get; set; }
        public bool RemovedFromSchedule { get; set; }
    }
}