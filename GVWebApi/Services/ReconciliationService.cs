using System.Collections.Generic;
using System.Linq;
using GV.CoFreedomDomain;
using GV.Domain;
using GV.Domain.Entities;
using GV.Domain.Views;
using GV.ExtensionMethods;
using GV.Services;
using GVWebapi.Models.Reconciliation;

namespace GVWebapi.Services
{
    public interface IReconciliationService
    {
        ReconciliationViewModel GetReconciliation(long cycleId);
        void UpdateReconCredit(ReconCreditSaveModel model);
    }

    public class ReconciliationService : IReconciliationService
    {
        private readonly IRepository _repository;
        private readonly ICoFreedomRepository _coFreedomRepository;
        private readonly IDeviceService _deviceService;
        private readonly ICoFreedomDeviceService _coFreedomDeviceService;
        private readonly IUnitOfWork _unitOfWork;

        public ReconciliationService(IRepository repository, ICoFreedomRepository coFreedomRepository, IDeviceService deviceService, ICoFreedomDeviceService coFreedomDeviceService, IUnitOfWork unitOfWork)
        {
            _repository = repository;
            _coFreedomRepository = coFreedomRepository;
            _deviceService = deviceService;
            _coFreedomDeviceService = coFreedomDeviceService;
            _unitOfWork = unitOfWork;
        }

        public ReconciliationViewModel GetReconciliation(long cycleId)
        {
            var cycle = _repository.Get<CyclesEntity>(cycleId);
            var model = new ReconciliationViewModel();
            SetDates(model, cycle);
            model.CostByDevice = GetCostByDevice(model, cycle);
            model.InvoicedService = GetInvoicedServices(model, cycle);
            model.CycleSummary = GetCycleSummary(cycle);
            return model;
        }

        private IList<CycleSummaryModel> GetCycleSummary(CyclesEntity cycle)
        {
            var cycleSummaryItems = cycle.CyclePeriods
                .Where(x => x.IsDeleted == false)
                .ToList();

            var cycleSummaryModels = new List<CycleSummaryModel>();

            foreach (var cyclePeriod in cycleSummaryItems)
            {
                var schedulesInThisPeriod = cyclePeriod.PeriodSchedules.Select(x => x.Schedule).ToList();
                
                var model = new CycleSummaryModel();
                model.Period = cyclePeriod.Period;
                model.AllocatedHardware = schedulesInThisPeriod.Sum(x => x.MonthlyHwCost);
                model.UnallocatedService = schedulesInThisPeriod.Sum(x => x.MonthlySvcCost);

                cycleSummaryModels.Add(model);
            }

            return cycleSummaryModels;
        }

        public void UpdateReconCredit(ReconCreditSaveModel model)
        {
            var cycleRecon = _repository.Get<CycleReconciliationServicesEntity>(model.CycleReconciliationServiceId);
            cycleRecon.Credit = model.Credit;
        }

        private IList<InvoicedServiceModel> GetInvoicedServices(ReconciliationViewModel model, CyclesEntity cycle)
        {
            var query = $"select * from _custom_placeholder.vw_csquarterlyhistory where customerid = {cycle.CustomerId} and date >= '{model.StartDate:yyyy/MM/dd}' and date <= '{model.EndDate:yyyy/MM/dd}'";
            var items = _coFreedomRepository.ExecuteSQL<ViewCsQuarterlyHistory>(query);

            var distinctMeterGroups = items.Select(x => x.ContractMeterGroup).Distinct().ToList();

            var invoicedServices = new List<InvoicedServiceModel>();

            foreach (var meterGroup in distinctMeterGroups)
            {
                var allItemsInThisMeterGroup = items.Where(x => x.ContractMeterGroup == meterGroup).ToList();
                var cycleRecon = cycle.ReconServices.FirstOrDefault(x => x.MeterGroup == meterGroup);
                if (cycleRecon == null)
                {
                    cycleRecon = new CycleReconciliationServicesEntity(meterGroup);
                    cycle.AddNewReconService(cycleRecon);
                    _unitOfWork.Commit();
                }

                var invoiceServiceModel = new InvoicedServiceModel();
                invoiceServiceModel.MeterGroup = meterGroup;
                invoiceServiceModel.ActualPages = allItemsInThisMeterGroup.Sum(x => x.CountedCopies);
                invoiceServiceModel.ContractedPages = allItemsInThisMeterGroup.Sum(x => x.CoveredCopies);
                invoiceServiceModel.OverageCost = allItemsInThisMeterGroup.Sum(x => x.TotalChargeAmount);
                invoiceServiceModel.Credit = cycleRecon.Credit;
                invoiceServiceModel.CycleReconciliationServiceId = cycleRecon.CycleReconciliationServiceId;
                invoicedServices.Add(invoiceServiceModel);
            }

            return invoicedServices;
        }

        private IList<CostByDeviceModel> GetCostByDevice(ReconciliationViewModel model, CyclesEntity cycle)
        {
            var query = $"select * from _custom_eviews.vw_monthly_device_costs where customerId = {cycle.CustomerId} and beginmeterdate >= '{model.StartDate:yyyy/MM/dd}' and endmeterdate <= '{model.EndDate:yyyy/MM/dd}'";
            var items = _coFreedomRepository.ExecuteSQL<ViewMonthlyDeviceCosts>(query);

            var costByDevices = new List<CostByDeviceModel>();
            var coFreedomDevices = _coFreedomDeviceService.GetCoFreedomDevices(cycle.CustomerId);

            var distinctEquipmentItems = items.Select(x => x.EquipmentNumber).Distinct().ToList();
            foreach (var equipmentNumber in distinctEquipmentItems)
            {
                var globalViewDevice = _deviceService.GetDevice(equipmentNumber);
                if(globalViewDevice == null) continue;

                var coFreedomDevice = coFreedomDevices.FirstOrDefault(x => x.EquipmentID == globalViewDevice.EquipmentId);
                
                var costByDeviceModel = new CostByDeviceModel();
                costByDeviceModel.SerialNumber = globalViewDevice.SerialNumber;
                costByDeviceModel.Model = globalViewDevice.Model;
                costByDeviceModel.Schedule = coFreedomDevice?.ScheduleNumber;
                costByDeviceModel.Location = coFreedomDevice?.Location;
                costByDeviceModel.User = coFreedomDevice?.AssetUser;
                costByDeviceModel.CostCenter = coFreedomDevice?.CostCenter;
                costByDeviceModel.BWCopies = GetBwCopies(items, equipmentNumber);
                costByDeviceModel.BWPrints = GetBwPrints(items, equipmentNumber);
                costByDeviceModel.ColorCopies = GetColorCopies(items, equipmentNumber);
                costByDevices.Add(costByDeviceModel);
            }

            return costByDevices;
        }

        private static decimal GetColorCopies(IList<ViewMonthlyDeviceCosts> items, string equipmentNumber)
        {
            return items
                .Where(x => x.EquipmentNumber.EqualsIgnore(equipmentNumber))
                .Where(x => x.ContractMeterGroup.EqualsIgnore("color laser prints")|| x.ContractMeterGroup.EqualsIgnore("color copies (color pages)"))
                .Sum(x => x.BilledAmount);
        }

        private static decimal GetBwPrints(IList<ViewMonthlyDeviceCosts> items, string equipmentNumber)
        {
            return items
                .Where(x => x.EquipmentNumber.EqualsIgnore(equipmentNumber))
                .Where(x => x.ContractMeterGroup.EqualsIgnore("b/w copies"))
                .Sum(x => x.BilledAmount);
        }

        private static decimal GetBwCopies(IList<ViewMonthlyDeviceCosts> items, string equipmentNumber)
        {
            return items
                .Where(x => x.EquipmentNumber.EqualsIgnore(equipmentNumber))
                .Where(x => x.ContractMeterGroup.EqualsIgnore("b/w laser prints"))
                .Sum(x => x.BilledAmount);
        }

        private static void SetDates(ReconciliationViewModel model, CyclesEntity cycle)
        {
            model.StartDate = cycle.CyclePeriods
                .Where(x => x.IsDeleted == false)
                .Select(x => x.Period)
                .Min();

            model.EndDate = cycle.CyclePeriods
                .Where(x => x.IsDeleted == false)
                .Select(x => x.Period)
                .Max();
        }
    }
}