using System;
using System.Collections.Generic;
using System.Linq;
using FluentDateTime;
using GV.CoFreedomDomain;
using GV.CoFreedomDomain.Entities;
using GV.Domain;
using GV.Domain.Entities;
using GV.Services;
using GVWebapi.Models.Devices;
using GVWebapi.Models.Schedules;
using GVWebapi.RemoteData;

namespace GVWebapi.Services
{
    public interface ICyclePeriodService
    {
        DateTime? AddPeriodToCycle(long cycleId);
        IList<CyclePeriodModel> GetCyclePeriods(long cycleId);
        IList<CyclePeriodModel> GetCyclePeriodsByPeriodId(long cyclePeriodId);
        DateTime? RemovePeriod(long cyclePeriodId);
        CyclePeriodSummaryModel GetCyclePeriodSummary(long cyclePeriodId);
        CyclePeriodSummaryModel RefreshCyclePeriodDevices(long cyclePeriodId);
        void SaveInstancesInvoiced(InvoiceInstanceSaveModel model);
        void UpdateInvoiceNumber(InvoiceNumberSaveModel model);
    }

    public class CyclePeriodService : ICyclePeriodService
    {
        private readonly IRepository _repository;
        private readonly IDeviceService _deviceService;
        private readonly IUnitOfWork _unitOfWork;
        private readonly ILocationsService _locationsService;
        private readonly ICoFreedomRepository _coFreedomRepository;
        private readonly IScheduleService _scheduleService;
        public CyclePeriodService(IRepository repository, IDeviceService deviceService, IUnitOfWork unitOfWork, ILocationsService locationsService, ICoFreedomRepository coFreedomRepository,IScheduleService scheduleService)
        {
            _repository = repository;
            _deviceService = deviceService;
            _unitOfWork = unitOfWork;
            _locationsService = locationsService;
            _coFreedomRepository = coFreedomRepository;
            _scheduleService = scheduleService;
        }

        public DateTime? AddPeriodToCycle(long cycleId)
        {
            var cycle = _repository.Get<CyclesEntity>(cycleId);
            var period = new CyclePeriodEntity(cycle);
         
            if (cycle.EndDate.HasValue == false)
            {
                var periodDate = cycle.StartDate;
                cycle.EndDate = periodDate;
                period.Period = periodDate;
            }
            else
            {
                var maxDate = cycle.CyclePeriods
                    .Where(x => x.IsDeleted == false)
                    .Max(x => x.Period);
                var periodDate = maxDate.AddMonths(1);
                cycle.EndDate = periodDate;
                period.Period = periodDate;
            }

            cycle.ModifiedDateTime = DateTimeOffset.Now;
            cycle.AddNewPeriod(period);
            return cycle.EndDate;
        }

        public IList<CyclePeriodModel> GetCyclePeriods(long cycleId)
        {
            var cyclePeriods = _repository.Find<CyclePeriodEntity>()
                .Where(x => x.Cycle.CycleId == cycleId)
                .Where(x => x.IsDeleted == false)
                .OrderBy(x => x.Period)
                .Select(x => new CyclePeriodModel
                {
                    CyclePeriodId = x.CyclePeriodId,
                    Period = x.Period,
                    Billed =0,
                    Allocated = 0,
                }).ToList();
              foreach(var cyclePeriod in cyclePeriods)
              {
                var _cyclePeriod = _repository.Get<CyclePeriodEntity>(cyclePeriod.CyclePeriodId);
                var _schudules = LoadScheduleServices(_cyclePeriod).ToList();
                cyclePeriod.Billed = _schudules.Sum(x => x.Total);
                cyclePeriod.Allocated = _schudules.Sum(x => x.MonthlyContractCost);

            }
                
          
            return cyclePeriods;
        }

        public IList<CyclePeriodModel> GetCyclePeriodsByPeriodId(long cyclePeriodId)
        {
            var cycle = _repository.Get<CyclePeriodEntity>(cyclePeriodId);
            return GetCyclePeriods(cycle.Cycle.CycleId);
        }

        public DateTime? RemovePeriod(long cyclePeriodId)
        {
            var cyclePeriod = _repository.Get<CyclePeriodEntity>(cyclePeriodId);
            cyclePeriod.IsDeleted = true;

            if (cyclePeriod.Cycle.EndDate == cyclePeriod.Period)
            {
                var cycle = _repository.Get<CyclesEntity>(cyclePeriod.Cycle.CycleId);

                var periods = cycle.CyclePeriods.Where(x => x.IsDeleted == false);
                if (periods.Any())
                {
                    var maxDate = cycle.CyclePeriods
                        .Where(x => x.IsDeleted == false)
                        .Max(x => x.Period);
                    cycle.EndDate = maxDate;
                }
                else
                {
                    cycle.EndDate = null;
                }

                return cycle.EndDate;
            }

            return cyclePeriod.Cycle.EndDate;
        }

        public CyclePeriodSummaryModel GetCyclePeriodSummary(long cyclePeriodId)
        {
            var cyclePeriod = _repository.Get<CyclePeriodEntity>(cyclePeriodId);
            var model = new CyclePeriodSummaryModel();
            model.PeriodDate = cyclePeriod.Period;
            model.Periods = GetCyclePeriodsByPeriodId(cyclePeriodId);
            model.InvoiceNumber = cyclePeriod.InvoiceNumber;
            model.Schedules = LoadScheduleServices(cyclePeriod);
            model.Devices = GetDevices(model.Schedules, cyclePeriod);
            return model;
        }

        public CyclePeriodSummaryModel RefreshCyclePeriodDevices(long cyclePeriodId)
        {
            var cyclePeriod = _repository.Get<CyclePeriodEntity>(cyclePeriodId);
            var model = new CyclePeriodSummaryModel(); 
            model.Schedules = LoadScheduleServices(cyclePeriod);
            model.Devices = GetDevices(model.Schedules, cyclePeriod);
            return model;
        }

        private List<DeviceModel> GetDevices(IList<CyclePeriodScheduleModel> modelSchedules, CyclePeriodEntity cyclePeriod)
        {
            var allDevices = new List<DeviceModel>();

            var minDate = cyclePeriod.Cycle.CyclePeriods.Min(x => x.Period).FirstDayOfMonth().Date;
            var maxDate = cyclePeriod.Cycle.CyclePeriods.Max(x => x.Period).LastDayOfMonth().Date;

            var deviceRates = _coFreedomRepository.Find<ViewEquipmentAndRate>()
                .Where(x => x.StartDate >= minDate)
                .Where(x => x.EndDate <= maxDate)
                .ToList();
            foreach (var schedule in modelSchedules)
            {
                var activeDevices = _deviceService.GetActiveDevices(schedule.ScheduleId);
                
                allDevices.AddRange(activeDevices);
            }


            return allDevices;
        }

        public IList<SchedulesModel> LoadSchedules(long customerId)
        {
            var schedules = _scheduleService.GetAcitveSchedulesByCustomer(customerId).ToList();
            return schedules;
        }

        public void SaveInstancesInvoiced(InvoiceInstanceSaveModel model)
        {
            var cyclePeriodSchedule = _repository.Get<CyclePeriodSchedulesEntity>(model.CyclePeriodScheduleId);
            cyclePeriodSchedule.InstancesInvoiced = model.InstancesInvoiced;
        }

        public void UpdateInvoiceNumber(InvoiceNumberSaveModel model)
        {
            var cyclePeriod = _repository.Get<CyclePeriodEntity>(model.CyclePeriodId);
            cyclePeriod.InvoiceNumber = model.InvoiceNumber;
            cyclePeriod.ModifiedDateTime = DateTimeOffset.Now;
        }

        public IList<CyclePeriodScheduleModel> LoadScheduleServices(CyclePeriodEntity cyclePeriod)
        {
            var cycle = cyclePeriod.Cycle;
            var schedules = _repository.Find<SchedulesEntity>()
                .Where(x => x.IsDeleted == false)
                .Where(x => x.EffectiveDateTime != null)
                .Where(x => x.ExpiredDateTime != null)
                .Where(x => x.CustomerId == cycle.CustomerId)
                .Where(x => cyclePeriod.Period >= x.EffectiveDateTime.Value.Date)
                .Where(x => cycle.EndDate.Value <= x.ExpiredDateTime.Value.Date)

                .Select(x => new CyclePeriodScheduleModel
                {
                    ScheduleId = x.ScheduleId,
                    ScheduleName = x.Name,
                    Service = x.MonthlySvcCost,
                    Hardware = x.MonthlyHwCost,                    
                    MonthlyContractCost = x.MonthlyContractCost,
 
                }).ToList();

            foreach (var schedule in schedules)
            {
                var cyclePeriodSchedule = cyclePeriod.PeriodSchedules.FirstOrDefault(x => x.Schedule.ScheduleId == schedule.ScheduleId);
                if (cyclePeriodSchedule == null)
                {
                    cyclePeriodSchedule = new CyclePeriodSchedulesEntity();
                    cyclePeriodSchedule.Schedule = _repository.Load<SchedulesEntity>(schedule.ScheduleId);
                    cyclePeriodSchedule.InstancesInvoiced = 1;
                    cyclePeriod.AddPeriodSchedule(cyclePeriodSchedule);
                    _unitOfWork.Commit();
                }

                schedule.InstancesInvoiced = cyclePeriodSchedule.InstancesInvoiced;
                schedule.CyclePeriodScheduleId = cyclePeriodSchedule.CyclePeriodScheduleId;
            }

            return schedules;
        }
    }

    public class InvoiceNumberSaveModel
    {
        public long CyclePeriodId { get; set; }
        public string InvoiceNumber { get; set; }
    }

    public class InvoiceInstanceSaveModel
    {
        public long CyclePeriodScheduleId { get; set; }
        public decimal InstancesInvoiced { get; set; }
    }

    public class CyclePeriodModel
    {
        public long CyclePeriodId { get; set; }
        public DateTime Period { get; set; }
        public decimal Billed { get; set; } 
        public decimal Allocated { get; set; }
    }

    public class CycleReconcilieModel
    {
        public long CyclePeriodId { get; set; }
        public bool IsReconciled { get; set; }
        public decimal Billed { get; set; }
        public decimal Allocated { get; set; }
    }

    public class CyclePeriodSummaryModel
    {
        public DateTime PeriodDate { get; set; }
        public string InvoiceNumber { get; set; }
     
        public IList<CyclePeriodModel> Periods { get; set; } = new List<CyclePeriodModel>();
        public IList<CyclePeriodScheduleModel> Schedules { get; set; } = new List<CyclePeriodScheduleModel>();
        public List<DeviceModel> Devices { get; set; } = new List<DeviceModel>();
    }

    public class CyclePeriodScheduleModel
    {
        public long CyclePeriodScheduleId { get; set; }  
        public long ScheduleId { get; set; }
        public string ScheduleName { get; set; }
        public decimal Service { get; set; }
        public decimal Hardware { get; set; }
        public decimal MonthlyContractCost { get; set; }
        public decimal InstancesInvoiced { get; set; }
        //used in the UI
        public decimal Total => Service + Hardware;
    }
}