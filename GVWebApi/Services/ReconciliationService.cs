using System.Collections.Generic;
using System.Linq;
using GV.CoFreedomDomain;
using GV.Domain;
using GV.Domain.Entities;
using GV.Domain.Views;
using GVWebapi.RemoteData;
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
            var PeriodDate = model.EndDate.AddMonths(1).AddDays(-1);
            var query = $"select * from _custom_placeholder.vw_REVisionInvoices where CustomerID = {cycle.CustomerId} and PeriodDate = '{PeriodDate:yyyy-MM-dd}'";
            //[vw_REVisionInvoices]
            var invoice = _coFreedomRepository.ExecuteSQL<vw_REVisionInvoices>(query).FirstOrDefault();

            var query1 = $"select * from _custom_placeholder.vw_RevisionHistory where InvoiceID = {invoice.InvoiceID} and OverageToDate = '{PeriodDate:yyyy-MM-dd}'";
            var items = _coFreedomRepository.ExecuteSQL<vw_RevisionHistory>(query1);

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
               var CPP =  allItemsInThisMeterGroup.Sum(x => x.CPP).Value;
                var invoiceServiceModel = new InvoicedServiceModel();
                invoiceServiceModel.MeterGroup = meterGroup;
                invoiceServiceModel.ActualPages = allItemsInThisMeterGroup.Sum(x => x.ActualVolume).Value;
                invoiceServiceModel.ContractedPages = allItemsInThisMeterGroup.Sum(x => x.ContractVolume).Value;
                invoiceServiceModel.BaseServiceForCycle = invoiceServiceModel.ContractedPages * CPP;
                var Overage = (invoiceServiceModel.ActualPages - invoiceServiceModel.ContractedPages);
                invoiceServiceModel.OverageCost = Overage > 0 ? Overage * CPP : 0;
                invoiceServiceModel.Credit = cycleRecon.Credit;
                invoiceServiceModel.CycleReconciliationServiceId = cycleRecon.CycleReconciliationServiceId;
                invoicedServices.Add(invoiceServiceModel);
            }

            return invoicedServices;
        }

        private IList<CostByDeviceModel> GetCostByDevice(ReconciliationViewModel model, CyclesEntity cycle)
        {
            var PeriodDate = model.EndDate.AddMonths(1).AddDays(-1);
            var query = $"select * from _custom_placeholder.vw_REVisionInvoices where CustomerID = {cycle.CustomerId} and PeriodDate = '{PeriodDate:yyyy-MM-dd}'";
            //[vw_REVisionInvoices]
            var invoice = _coFreedomRepository.ExecuteSQL<vw_REVisionInvoices>(query).FirstOrDefault();

            var query1 = $"select * from _custom_eviews.vw_monthly_device_costs where InvoiceID = {invoice.InvoiceID}";
            var items = _coFreedomRepository.ExecuteSQL<ViewMonthlyDeviceCosts>(query1);

            var costByDevices = new List<CostByDeviceModel>();
            var coFreedomDevices = _coFreedomDeviceService.GetCoFreedomDevices(cycle.CustomerId);

            var distinctEquipmentItems = items.Select(x => x.EquipmentID).Distinct().ToList();
            foreach (var equipmentID in distinctEquipmentItems)
            {
                var globalViewDevice = _deviceService.GetDevice(equipmentID);
                if(globalViewDevice == null) continue;

                var coFreedomDevice = coFreedomDevices.FirstOrDefault(x => x.EquipmentID == globalViewDevice.EquipmentID);
                
                var costByDeviceModel = new CostByDeviceModel();
                costByDeviceModel.SerialNumber = globalViewDevice.SerialNumber;
                costByDeviceModel.Model = globalViewDevice.Model;
                costByDeviceModel.Schedule = coFreedomDevice?.ScheduleNumber;
                costByDeviceModel.Location = coFreedomDevice?.Location;
                costByDeviceModel.User = coFreedomDevice?.AssetUser;
                costByDeviceModel.CostCenter = coFreedomDevice?.CostCenter;
                costByDeviceModel.BWCopies = GetBwCopies(items, equipmentID);
                costByDeviceModel.BWPrints = GetBwPrints(items, equipmentID);
                costByDeviceModel.ColorCopies = GetColorCopies(items, equipmentID);
                costByDevices.Add(costByDeviceModel);
            }

            return costByDevices;
        }

        private static decimal GetColorCopies(IList<ViewMonthlyDeviceCosts> items, int equipmentNumber)
        {
            return items
                .Where(x => x.EquipmentID == equipmentNumber)
                .Where(x => x.ContractMeterGroup.EqualsIgnore("color laser prints")|| x.ContractMeterGroup.EqualsIgnore("color copies (color pages)"))
                .Sum(x => x.BilledAmount);
        }

        private static decimal GetBwPrints(IList<ViewMonthlyDeviceCosts> items, int equipmentNumber)
        {
            return items
                .Where(x => x.EquipmentID == equipmentNumber)
                .Where(x => x.ContractMeterGroup.EqualsIgnore("b/w copies"))
                .Sum(x => x.BilledAmount);
        }

        private static decimal GetBwCopies(IList<ViewMonthlyDeviceCosts> items, int equipmentNumber)
        {
            return items
                .Where(x => x.EquipmentID == equipmentNumber)
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