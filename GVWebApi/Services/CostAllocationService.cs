using System;
using System.Collections.Generic;
using System.Linq;
using GVWebapi.Models.CostAllocation;
using GVWebapi.Models.Schedules;
using GVWebapi.RemoteData;
using GV.CoFreedomDomain;
using GV.CoFreedomDomain.Entities;
using GV.Domain;
using GVWebapi.Models.Devices;
using GV.Domain.Entities;

namespace GVWebapi.Services
{
    public interface ICostAllocationService
    {
     
        IEnumerable<CostAllocationMeterGroup> GetCostAllocationMeterGroups(int custid);
    
        
        IEnumerable<CostCenterModel> GetCostCenters(CyclePeriodScheduleModel modelSchedule);
        IEnumerable<ServiceCostCenterViewModel> GetScheduleCostCenters(long scheduleId, List<ScheduleCostCenter> scheduleCostCenters);
        List<AllocatedServicesViewModel> GetScheduleAllocatedServices(List<CostCenterModel> models, List<MeterGroup> MeterGroups, decimal taxRate);
        List<CostCenterSummaryViewModel> GetCostCeterSummaryServices(List<CostCenterModel> models, List<ScheduleDevicesModel> deviceModels, List<MeterGroup> MeterGroups, decimal taxRate);
        decimal GetAllocatedServicesTotal(CyclePeriodEntity modelSchedule);
    }

    public class CostAllocationService : ICostAllocationService
    {
        private readonly ICoFreedomRepository _coFreedomRepository;
        private readonly IRepository _repository;
        private readonly IUnitOfWork _unitOfWork;
        private readonly IScheduleServicesService _ischedueService;
        private readonly ICoFreedomDeviceService _deviceService;

        public CostAllocationService(IRepository repository, ICoFreedomRepository coFreedomRepository, IUnitOfWork unitOfWork, IScheduleServicesService ischeduleservice, ICoFreedomDeviceService deviceService)
        {
            _repository = repository;
            _coFreedomRepository = coFreedomRepository;
            _unitOfWork = unitOfWork;
            _ischedueService = ischeduleservice;
            _deviceService = deviceService;

        }
 
        // Function being Used 2019-JULY-2
        public IEnumerable<CostAllocationMeterGroup> GetCostAllocationMeterGroups(int custid)
        {
            try
            {
                GlobalViewEntities db = new GlobalViewEntities();
                var metergroups = GetCoFreedomMeterGroups(custid);
               
                foreach (var metergroup in metergroups)
                {
                    var _metergroupexists = db.CostAllocationMeterGroups.FirstOrDefault(x => x.ContractMeterGroupID == metergroup.ContractMeterGroupID && x.CustomerID == custid);
                    if (_metergroupexists == null)
                    {
                        var _metergroup = new CostAllocationMeterGroup();
                        _metergroup.ContractMeterGroupID = metergroup.ContractMeterGroupID;
                        _metergroup.ContractMeterGroup = metergroup.ContractMeterGroup;
                        _metergroup.ExcessCPP = 0.00M;
                        _metergroup.MeterGroupID = db.RevisionMeterGroups.FirstOrDefault(x => x.ERPMeterGroupID == metergroup.ContractMeterGroupID).MeterGroupID;
                        _metergroup.CustomerID = custid;
                        db.CostAllocationMeterGroups.Add(_metergroup);
                    }
    
                }
                db.SaveChanges();
                var _settingsMeterGroup = db.CostAllocationMeterGroups.Where(x => x.CustomerID == custid).ToList();
                

              return _settingsMeterGroup;

            } catch(Exception ex)
            {
                return null;
            }
        }
        // Function being Used 2019-JULY-2
        public IEnumerable<CostCenterModel> GetCostCenters(CyclePeriodScheduleModel modelSchedule)
        {

            try
            {
                var  CostCenters = new List<CostCenterModel>();
                GlobalViewEntities db = new GlobalViewEntities();
                var CostCenterMeterGroups = db.ScheduleCostCenters.Where(x => x.ScheduleID == modelSchedule.ScheduleId).ToList();
              
                foreach(var costcenter in CostCenterMeterGroups)
                {
                    var settings = db.ScheduleServices.Where(x => x.ContractMeterGroupID == costcenter.MeterGroupID.Value).FirstOrDefault();
                    var costCenter = new CostCenterModel();
                    costCenter.ScheduleID = costcenter.ScheduleID;
                    costCenter.ScheduleName = modelSchedule.ScheduleName;
                    costCenter.ContractMeterGroupID = costcenter.MeterGroupID.Value;
                    costCenter.MeterGroupDesc = settings.MeterGroup;
                    costCenter.CostCenter = costcenter.CostCenter;
                    costCenter.BaseCPP = settings.BaseCpp;
                    costCenter.Volume = costcenter.Volume;
                    costCenter.TaxRate = ServiceTaxRate(modelSchedule.ScheduleId);
                    costCenter.InstanceInvoiced = modelSchedule.InstancesInvoiced;
                    costCenter.Removed = RemovedFromSchedule(modelSchedule.ScheduleId, costcenter.MeterGroupID.Value);
                    CostCenters.Add(costCenter);
                }
                return CostCenters;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public decimal GetAllocatedServicesTotal(CyclePeriodEntity modelSchedule)
        {

            try
            {
                var CostCenters = new List<CostCenterModel>();
                var Schedules = modelSchedule.PeriodSchedules.Select(x => x.Schedule.ScheduleId).ToList();
                GlobalViewEntities db = new GlobalViewEntities();
                var CostCenterMeterGroups = db.ScheduleCostCenters.Where(x => Schedules.Contains(x.ScheduleID)).ToList();

                foreach (var costcenter in CostCenterMeterGroups)
                {
                    var settings = db.CostAllocationMeterGroups.Where(x => x.ContractMeterGroupID == costcenter.MeterGroupID.Value).FirstOrDefault();
                    var costCenter = new CostCenterModel();
                    costCenter.ExcessCPP = settings.ExcessCPP;
                    costCenter.Volume = costcenter.Volume;
                    CostCenters.Add(costCenter);
                }
                var total = CostCenters.Sum(x => x.Volume * x.ExcessCPP);
                return total.Value;
            }
            catch (Exception ex)
            {
                return 0.00M;
            }
        }
        // Function being Used 2019-JULY-2
        public IEnumerable<ServiceCostCenterViewModel> GetScheduleCostCenters(long scheduleId, List<ScheduleCostCenter> scheduleCostCenters)
        {

            try
            {
                GlobalViewEntities db = new GlobalViewEntities();
                var serviceCostCenters = new List<ServiceCostCenterViewModel>();
                var CostCenterMeterGroups = _ischedueService.GetMeterGroups(scheduleId).Where(x => x.RemovedFromSchedule == false).OrderBy(x => x.MeterGroup);
                var schedule = db.Schedules.Find(scheduleId);
                var devices = _deviceService.GetCoFreedomDevices(schedule.Name, schedule.CustomerId);
                var costCenters = new List<string>();
                if (devices.Count > 0)
                {
                     costCenters = devices.Select(x => x.CostCenter).Distinct().ToList();
                } else
                {
                    costCenters = GetCoFreedomCostCenters(schedule.CustomerId);
                }
              
               
                foreach (var costcenter in costCenters)
                {
                    var serviceCostCenter = new ServiceCostCenterViewModel();
                    serviceCostCenter.ScheduleID = scheduleId;
                    serviceCostCenter.CustomerID = schedule.CustomerId;
                    serviceCostCenter.CostCenter = costcenter;
                    foreach (var meterGroup in CostCenterMeterGroups)
                    {
                        var scheduleCostCenter = scheduleCostCenters.Where(x => x.MeterGroupID == meterGroup.ContractMeterGroupID && x.CostCenter == costcenter).FirstOrDefault();
                        var MeterGroup = new MeterGroupCostCenter();
                        MeterGroup.ContractMeterGroupID = meterGroup.ContractMeterGroupID;
                        MeterGroup.MeterGroupDesc = meterGroup.MeterGroup;
                        MeterGroup.Volume = scheduleCostCenter == null ? 0 : scheduleCostCenter.Volume.Value;
                        serviceCostCenter.MeterGroups.Add(MeterGroup);
                    }
                    serviceCostCenters.Add(serviceCostCenter);
                }
                return serviceCostCenters.OrderBy(x => x.CostCenter);
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        // Function being Used 2019-JULY-2 NOTE Fix this to use CyclePeriodId
        private decimal InstanceMultipler(long scheduleId)
        {
            using (var context = new GlobalViewEntities())
            {
                var cycleperiod = context.CyclePeriodSchedules.Where(x => x.ScheduleId == scheduleId).OrderByDescending(x => x.CyclePeriodScheduleId).FirstOrDefault();
                return cycleperiod.InstancesInvoiced;
            }

        }
        // Function being Used 2019-JULY-2
        private decimal  ServiceTaxRate(long scheduleId)
        {
            using (var context = new GlobalViewEntities())
            {
                var custid = context.Schedules.Find(scheduleId).CustomerId;
                var taxrate = context.Locations.Where(x => x.CustomerId == custid && x.IsCorporate == true).Select(x => x.TaxRate).FirstOrDefault();
                return taxrate;
            }
        }
        // Function being Used 2019-JULY-2
        private bool RemovedFromSchedule(long scheduleId, long MeterGroupId)
        {
            using (var context = new GlobalViewEntities())
            {
               
                var scheduleService = context.ScheduleServices.FirstOrDefault(x => x.ScheduleId == scheduleId && x.ContractMeterGroupID == MeterGroupId);
                return scheduleService.RemovedFromSchedule;
            }
        }
        // Function being Used 2019-JULY-2
        private IEnumerable<MeterGroup> GetCoFreedomMeterGroups(int customerId)
        {
            return _coFreedomRepository.Find<ScEquipmentEntity>()
                .Where(x => x.CustomerId == customerId)
                .SelectMany(x => x.ContractDetails)
                .SelectMany(x => x.Contract.ContractMeterGroups)
                  .Select(x => new MeterGroup { ContractMeterGroupID = x.ContractMeterGroupId, ContractMeterGroup = x.ContractMeterGroup })
                .Distinct()
                .ToList();
        }
        // Function being Used 2019-JULY-2
        private List<string> GetCoFreedomCostCenters(long customerId)
        {
            return _coFreedomRepository.Find<ScEquipmentEntity>()
                .Where(x => x.CustomerId == customerId)
                .SelectMany(x => x.CustomProperties)
                .Where(x => x.ShAttributeId ==  2001)
                .Select(x => x.TextVal)
                .Distinct()
                .ToList();
        }
        // Function being Used 2019-JULY-2
        public List<AllocatedServicesViewModel> GetScheduleAllocatedServices(List<CostCenterModel> models, List<MeterGroup> MeterGroups, decimal taxRate)
        {
            var allocatedServices =  new List<AllocatedServicesViewModel>();
            var costCenters = models.Select(x => new { costcenter = x.CostCenter, schedulename = x.ScheduleName, invoiceInstance = x.InstanceInvoiced }).Distinct();
            foreach(var costCenter in costCenters)
            {
                AllocatedServicesViewModel service = new AllocatedServicesViewModel();
                service.ScheduleName = costCenter.schedulename;
                service.CostCenter = costCenter.costcenter;
                var model = models.Where(x => x.ScheduleName == costCenter.schedulename && x.CostCenter == costCenter.costcenter);

                foreach (var metergoup in MeterGroups)
                {
                    var Row = model.Where(x => x.ContractMeterGroupID == metergoup.ContractMeterGroupID).FirstOrDefault();
                    MeterGroupCostCenter mg = new MeterGroupCostCenter();
                    mg.MeterGroupDesc = metergoup.ContractMeterGroup;
                    mg.Volume = Row == null ? 0 : Row.Volume;
                    service.MeterGroups.Add(mg);
                }
                
                
                service.ServiceCost = model.Sum(x => x.Volume.Value * x.BaseCPP.Value);
                service.ServiceCostTax = service.ServiceCost * (taxRate /100);
                service.TotalCost = service.ServiceCost + service.ServiceCostTax;
                service.InstanceInvoiced = costCenter.invoiceInstance; 
                allocatedServices.Add(service);
            }

            return allocatedServices;
        }
        public List<CostCenterSummaryViewModel> GetCostCeterSummaryServices(List<CostCenterModel> models, List<ScheduleDevicesModel> deviceModels, List<MeterGroup> MeterGroups, decimal taxRate)
        {
            var summaries = new List<CostCenterSummaryViewModel>();
            var allocatedServices = new List<AllocatedServicesViewModel>();
            var costCenters = models.Select(x => new { costcenter = x.CostCenter, invoiceInstance = x.InstanceInvoiced }).Distinct();
            foreach (var costCenter in costCenters)
            {
                CostCenterSummaryViewModel summary = new CostCenterSummaryViewModel();
               
                summary.CostCenter = costCenter.costcenter;
                var model = models.Where(x => x.CostCenter == costCenter.costcenter);

                foreach (var metergoup in MeterGroups)
                {
                    var Row = model.Where(x => x.ContractMeterGroupID == metergoup.ContractMeterGroupID).FirstOrDefault();
                    MeterGroupCostCenter mg = new MeterGroupCostCenter();
                    mg.MeterGroupDesc = metergoup.ContractMeterGroup;
                    mg.Volume = Row == null ? 0 : Row.Volume;
                    summary.MeterGroups.Add(mg);
                }
                summary.Hardware = deviceModels.Where(x => x.CostCenter == costCenter.costcenter).Sum(x => x.MonthlyCost);
                summary.HardwareTax = deviceModels.Where(x => x.CostCenter == costCenter.costcenter).Sum(x => x.CalculatedTax);
                summary.Service = model.Where(x => x.CostCenter == costCenter.costcenter).Sum(x => x.Volume.Value * x.BaseCPP.Value);
                summary.ServiceTax = taxRate;
                summary.Adjustments = 0.00M;
                summary.InstanceInvoiced = costCenter.invoiceInstance;
                summaries.Add(summary);
            }
            return summaries;
        }
        
      
    }

}