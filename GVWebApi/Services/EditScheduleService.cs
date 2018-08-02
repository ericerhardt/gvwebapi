using GV.Domain;
using GV.Domain.Entities;
using GVWebapi.Models.Schedules;

namespace GVWebapi.Services
{
    public interface IEditScheduleService
    {
        EditScheduleTopModel GetScheduleTopModel(long scheduleId);
    }

    public class EditScheduleService : IEditScheduleService
    {
        private readonly IRepository _repository;
        private readonly IScheduleService _scheduleService;
        private readonly ICoFreedomDeviceService _coFreedomDeviceService;

        public EditScheduleService(IRepository repository, IScheduleService scheduleService, ICoFreedomDeviceService coFreedomDeviceService)
        {
            _repository = repository;
            _scheduleService = scheduleService;
            _coFreedomDeviceService = coFreedomDeviceService;
        }

        public EditScheduleTopModel GetScheduleTopModel(long scheduleId)
        {
            var scheduleEntity = _repository.Get<SchedulesEntity>(scheduleId);
            var model = EditScheduleTopModel.For(scheduleEntity);
            model.SetSchedule(_scheduleService.GetAll(scheduleEntity.CustomerId));
            _coFreedomDeviceService.LoadCoFreedomDevices(scheduleId);
            return model;
        }
    }
}