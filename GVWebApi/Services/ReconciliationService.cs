using System.Collections.Generic;
using System.Linq;
using GV.CoFreedomDomain;
using GV.Domain;
using GV.Domain.Entities;
using GVWebapi.Entities;
using GV.Domain.Views;
using GVWebapi.RemoteData;
using GV.ExtensionMethods;
using GV.Services;
using GVWebapi.Models.Reconciliation;
using System;
using GVWebapi.Models.Schedules;
using GVWebapi.Models.CostAllocation;

namespace GVWebapi.Services
{
    public interface IReconciliationService
    {
        ReconciliationViewModel GetReconciliation(long cycleId, int custid);
        ReconciliationViewModel GetReconciliationSummary(long cycleId, int custid);
        void UpdateReconCredit(ReconCreditSaveModel model);
    }

    public class ReconciliationService : IReconciliationService
    {
        private readonly IRepository _repository;
        private readonly ICoFreedomRepository _coFreedomRepository;
        private readonly IScheduleDevicesService _deviceService;
        private readonly ICostAllocationService _allocatedServices;
        private readonly ICoFreedomDeviceService _coFreedomDeviceService;
        private readonly IScheduleService _scheduleService;
        private readonly ILocationsService _locationsService;
        private readonly ICyclePeriodService _cyclePeriodService;
        private readonly IUnitOfWork _unitOfWork;

        public ReconciliationService(IRepository repository, ICoFreedomRepository coFreedomRepository, IScheduleDevicesService deviceService, ICoFreedomDeviceService coFreedomDeviceService, IUnitOfWork unitOfWork,IScheduleService scheduleService, ICostAllocationService allocatedServices,ICyclePeriodService cyclePeriodService, ILocationsService locationsService)
        {
            _repository = repository;
            _coFreedomRepository = coFreedomRepository;
            _deviceService = deviceService;
            _allocatedServices = allocatedServices;
            _locationsService = locationsService;
            _cyclePeriodService = cyclePeriodService;
            _coFreedomDeviceService = coFreedomDeviceService;
            _scheduleService = scheduleService;
            _unitOfWork = unitOfWork;
        }

        public ReconciliationViewModel GetReconciliation(long cycleId, int custid)
        {
            var cycle = _repository.Get<CyclesEntity>(cycleId);
            GlobalViewEntities db = new GlobalViewEntities();
            var assumedAllocations = db.CostAllocationSettings.Where(x => x.CustomerID == custid).FirstOrDefault().AssumedAllocation;
            var taxrate = db.Locations.Where(x => x.CustomerId == custid && x.IsCorporate == true).FirstOrDefault().TaxRate;
            var model = new ReconciliationViewModel();
            SetDates(model, cycle);
            model.InvoicedService = GetInvoicedServices(model, cycle, taxrate);
            var schedules = GetScheduleCycle(cycle);
            var serviceModel = GetAllocatedCostCenters(schedules);
            var ReConcileSummary = GetCostByCostCenter(model, cycle, taxrate);
            model.CostByCostCenter = ReConcileSummary.summaries;
            model.MeterGroups = ReConcileSummary.MeterGroups;
            var hardwareTax = model.CostByCostCenter.Sum(x => x.HardwareTax);
            model.ReconcileAdj = cycle.ReconcileAdj;
            model.CycleSummary = GetCycleSummary(cycle, taxrate, hardwareTax);
            return model;
        }
        public ReconciliationViewModel GetReconciliationSummary(long cycleId, int custid)
        {
            GlobalViewEntities db = new GlobalViewEntities();
            var assumedAllocations = db.CostAllocationSettings.Where(x => x.CustomerID == custid).FirstOrDefault().AssumedAllocation;
            var taxrate = db.Locations.Where(x => x.CustomerId == custid && x.IsCorporate == true).FirstOrDefault().TaxRate;
            var cycle = _repository.Get<CyclesEntity>(cycleId);
            var model = new ReconciliationViewModel();
            SetDates(model, cycle);
            model.InvoicedService = GetInvoicedServices(model, cycle,taxrate);
  
            return model;
        }

        private IList<CycleSummaryModel> GetCycleSummary(CyclesEntity cycle, decimal taxrate, decimal hardwaretax)
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
                model.AllocatedHardwareTax = hardwaretax;
                model.AllocatedService = _allocatedServices.GetAllocatedServicesTotal(cyclePeriod);
                model.AllocatedServiceTax = taxrate;
                model.UnallocatedService = 0.00M;
                model.UnallocatedServiceTax = 0.00M;
                model.Adjustments = 0.00M;
                cycleSummaryModels.Add(model);
            }

            return cycleSummaryModels;
        }

        public List<CostCenterModel> GetAllocatedCostCenters(IList<CyclePeriodScheduleModel> modelSchedules)
        {
            var allCostcenters = new List<CostCenterModel>();

            foreach (var schedule in modelSchedules)
            {
                var costCenters = _allocatedServices.GetCostCenters(schedule);


                allCostcenters.AddRange(costCenters);
            }
            return allCostcenters;
        }
        public IList<CyclePeriodScheduleModel> GetScheduleCycle(CyclesEntity cycle)  
        {
            var schedules = cycle.CyclePeriods.SelectMany(x => x.PeriodSchedules).Select(x => x.Schedule)
           
            .Select(x => new CyclePeriodScheduleModel
             {
                 ScheduleId = x.ScheduleId,
                 ScheduleName = x.Name,
                 CustomerId = x.CustomerId,
                 MonthlyContractCost = x.MonthlyContractCost,

             }).ToList();

            return schedules;
        }
        private IList<CycleSummaryModel> GetCycleSummary2(CyclesEntity cycle)
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
        private IList<InvoicedServiceModel> GetInvoicedServices(ReconciliationViewModel model, CyclesEntity cycle, decimal taxrate)
        {
            var PeriodDate = model.EndDate.AddMonths(1).AddDays(-1);
            var query = $"select * from _custom_placeholder.vw_REVisionInvoices where CustomerID = {cycle.CustomerId} and PeriodDate = '{PeriodDate:yyyy-MM-dd}'";
            //[vw_REVisionInvoices]
            var invoice = _coFreedomRepository.ExecuteSQL<vw_REVisionInvoices>(query).FirstOrDefault();

            var query1 = $"select * from _custom_placeholder.vw_RevisionHistory where InvoiceID = {invoice.InvoiceID} and OverageToDate = '{PeriodDate:yyyy-MM-dd}'";
            var items = _coFreedomRepository.ExecuteSQL<vw_RevisionHistory>(query1);

            var basecpp = _scheduleService.GetCCSummary(cycle.CustomerId).ToList();

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
                var CPP = allItemsInThisMeterGroup.Sum(x => x.CPP).Value;
                var baseCPP = basecpp.Where(x => x.MeterGroup == meterGroup).FirstOrDefault();
                var invoiceServiceModel = new InvoicedServiceModel();
                invoiceServiceModel.MeterGroup = meterGroup;
                invoiceServiceModel.ActualPages = allItemsInThisMeterGroup.Sum(x => x.ActualVolume).Value;
                invoiceServiceModel.ContractedPages = allItemsInThisMeterGroup.Sum(x => x.ContractVolume).Value;
                invoiceServiceModel.BaseServiceForCycle = invoiceServiceModel.ContractedPages * baseCPP.BaseCPP;
                var Overage = (invoiceServiceModel.ActualPages - invoiceServiceModel.ContractedPages);
                invoiceServiceModel.OverageCost = Overage > 0 ? Overage * CPP : 0;
                invoiceServiceModel.TaxRate = taxrate;
                invoiceServiceModel.Credit = cycleRecon.Credit;
                invoiceServiceModel.CycleReconciliationServiceId = cycleRecon.CycleReconciliationServiceId;
                invoicedServices.Add(invoiceServiceModel);
            }

            return invoicedServices;
        }
        //private IList<InvoicedServiceModel> GetInvoicedServices(ReconciliationViewModel model, CyclesEntity cycle, bool AssumedVolumes,decimal taxrate)
        //{
        //    var PeriodDate = model.EndDate.AddMonths(1).AddDays(-1);
        //    var query = $"select * from _custom_placeholder.vw_REVisionInvoices where CustomerID = {cycle.CustomerId} and PeriodDate = '{PeriodDate:yyyy-MM-dd}'";
        //    //[vw_REVisionInvoices]
        //    var invoice = _coFreedomRepository.ExecuteSQL<vw_REVisionInvoices>(query).FirstOrDefault();

        //    var query1 = $"select * from _custom_placeholder.vw_RevisionHistory where InvoiceID = {invoice.InvoiceID} and OverageToDate = '{PeriodDate:yyyy-MM-dd}'";
        //    var items = _coFreedomRepository.ExecuteSQL<vw_RevisionHistory>(query1);

        //    var basecpp =  _scheduleService.GetCCSummary(cycle.CustomerId).ToList();

        //    var distinctMeterGroups = items.Select(x => x.ContractMeterGroup).Distinct().ToList();


        //    var invoicedServices = new List<InvoicedServiceModel>();

        //    foreach (var meterGroup in distinctMeterGroups)
        //    {
        //        var allItemsInThisMeterGroup = items.Where(x => x.ContractMeterGroup == meterGroup).ToList();
        //        var cycleRecon = cycle.ReconServices.FirstOrDefault(x => x.MeterGroup == meterGroup);
        //        if (cycleRecon == null)
        //        {
        //            cycleRecon = new CycleReconciliationServicesEntity(meterGroup);
        //            cycle.AddNewReconService(cycleRecon);
        //            _unitOfWork.Commit();
        //        }
        //       var CPP =  allItemsInThisMeterGroup.Sum(x => x.CPP).Value;
        //        var baseCPP = basecpp.Where(x => x.MeterGroup == meterGroup).FirstOrDefault();
        //        var invoiceServiceModel = new InvoicedServiceModel();
        //        invoiceServiceModel.MeterGroup = meterGroup;
        //        invoiceServiceModel.ActualPages = allItemsInThisMeterGroup.Sum(x => x.ActualVolume).Value;
        //        invoiceServiceModel.ContractedPages = allItemsInThisMeterGroup.Sum(x => x.ContractVolume).Value;
        //        if (AssumedVolumes)
        //        {
        //            invoiceServiceModel.BaseServiceForCycle = 0;
        //        } else
        //        {
        //            invoiceServiceModel.BaseServiceForCycle = invoiceServiceModel.ContractedPages * baseCPP.BaseCPP;
        //        }
                
        //        var Overage = (invoiceServiceModel.ActualPages - invoiceServiceModel.ContractedPages);
        //        invoiceServiceModel.OverageCost = Overage > 0 ? Overage * CPP : 0;
        //        invoiceServiceModel.Credit = cycleRecon.Credit;
        //        invoiceServiceModel.TaxRate = taxrate;
        //        invoiceServiceModel.CycleReconciliationServiceId = cycleRecon.CycleReconciliationServiceId;
        //        invoicedServices.Add(invoiceServiceModel);
        //    }

        //    return invoicedServices;
       // }

        private ReconcileCostCenterSummary GetCostByCostCenter(ReconciliationViewModel model, CyclesEntity cycle,  decimal taxrate)
        {
            ReconcileCostCenterSummary ReconcileSummary = new ReconcileCostCenterSummary();
         
            var PeriodDate = model.EndDate.AddMonths(1).AddDays(-1);
            var query = $"select * from _custom_placeholder.vw_REVisionInvoices where CustomerID = {cycle.CustomerId} and PeriodDate = '{PeriodDate:yyyy-MM-dd}'";
            //[vw_REVisionInvoices]
            var invoice = _coFreedomRepository.ExecuteSQL<vw_REVisionInvoices>(query).FirstOrDefault();

            try
            {
                var query1 = $"select * from vw_GV_InvoicedEquipmentHistory where InvoiceID = {invoice.InvoiceID} ";
                var invoicedEquipment = _coFreedomRepository.ExecuteSQL<vw_GV_InvoicedEquipmentHistory>(query1);

                var basecpp = _scheduleService.GetCCSummary(cycle.CustomerId).ToList();

                var distinctMeterGroups = invoicedEquipment.Select(x => x.ContractMeterGroup).Distinct().ToList();


 
            var items = new List<ReconcileCostCenter>();

            foreach(var device in invoicedEquipment)
             {
              
                var HardwareTax = _locationsService.GetTaxRate(device.Location);
                items.Add(ReconcileCostCenter.For(device, HardwareTax));

            }
            var costCenters = items.OrderBy(x => x.CostCenter).Select(x => x.CostCenter).Distinct().ToList();
            var meterGroups = _allocatedServices.GetCostAllocationMeterGroups((int)cycle.CustomerId);
            
            int i = 1;
            
            foreach(var metergroup in meterGroups)
                {
                    var group = new MeterGroup();
                    group.ContractMeterGroup = metergroup.ContractMeterGroup;
                    group.ContractMeterGroupID = metergroup.ContractMeterGroupID.Value;
                    ReconcileSummary.MeterGroups.Add(group);
                }
 
            foreach (var costcenter in costCenters)
            {
                var summary = new CostCenterSummaryViewModel();
                summary.RowNumber = i++;
                summary.CostCenter = costcenter;
                var item = items.Where(x => x.CostCenter == costcenter).ToList();
                foreach(var meterGroup in meterGroups)
                {
                        var volume = item.Where(x => x.ContractMeterGroupID  == meterGroup.ContractMeterGroupID).Sum(x => x.Volume);
                        if (volume != 0)
                        {
                            MeterGroupCostCenter mg = new MeterGroupCostCenter();
                            mg.MeterGroupDesc = meterGroup.ContractMeterGroup;
                            mg.Volume = volume;
                            mg.ExcessCPP = meterGroup.ExcessCPP;
                            summary.MeterGroups.Add(mg);
                        }
                        else
                        {
                            MeterGroupCostCenter mg = new MeterGroupCostCenter();
                            mg.MeterGroupDesc = meterGroup.ContractMeterGroup;
                            mg.Volume = 0;
                            summary.MeterGroups.Add(mg);
                        }
                }
                summary.Hardware = item.Where(x => x.CostCenter == costcenter).Sum(x => x.MonthlyCost);
                summary.HardwareTax = item.Where(x => x.CostCenter == costcenter).Sum(x => x.CalculatedTax);
                summary.Service =  summary.MeterGroups.Where(x => x.Volume != 0).Sum(x => x.Volume * x.ExcessCPP.Value).Value;
                summary.ServiceTax = summary.Service * (taxrate / 100);
                summary.Adjustments = 0.00M;
                ReconcileSummary.summaries.Add(summary);
            }

            return ReconcileSummary;

            }
            catch (Exception ex)
            {
                return null;
            }
        }

        private IList<CostByDeviceModel> GetCostByDevice(ReconciliationViewModel model, CyclesEntity cycle, IList<InvoicedServiceModel> services)
        {
            var PeriodDate = model.EndDate.AddMonths(1).AddDays(-1);
            var query = $"select * from _custom_placeholder.vw_REVisionInvoices where CustomerID = {cycle.CustomerId} and PeriodDate = '{PeriodDate:yyyy-MM-dd}'";
            //[vw_REVisionInvoices]
            var invoice = _coFreedomRepository.ExecuteSQL<vw_REVisionInvoices>(query).FirstOrDefault();

            var query1 = $"select * from _custom_eviews.vw_monthly_device_costs where InvoiceID = {invoice.InvoiceID}";
            var items = _coFreedomRepository.ExecuteSQL<ViewMonthlyDeviceCosts>(query1);

           // var query2 = $"select * from _custom_placeholder.vw_RevisionHistory where InvoiceID = {invoice.InvoiceID} and OverageToDate = '{PeriodDate:yyyy-MM-dd}'";
           // var items2 = _coFreedomRepository.ExecuteSQL<vw_RevisionHistory>(query2);
            Dictionary<string , DeviceCostsModel> contractedVolume = new Dictionary<string, DeviceCostsModel>();
            foreach (var item in services)
            {
                DeviceCostsModel devicecost = new DeviceCostsModel();
                devicecost.Cost = item.BaseServiceForCycle;
                devicecost.Overage = item.OverageCost;
                contractedVolume.Add(item.MeterGroup, devicecost);
            }
            var costByDevices = new List<CostByDeviceModel>();
            var coFreedomDevices = _coFreedomDeviceService.GetCoFreedomDevices(cycle.CustomerId);
            var adj = 0.0000M;
            var bw = items.Where(x => x.MeterType.EqualsIgnore("B\\W")).Count();
            var clr = items.Where(x => x.MeterType.EqualsIgnore("Color")).Count();
            var distinctEquipmentItems = items.Select(x => x.EquipmentID).Distinct().ToList();
            if (cycle.ReconcileAdj > 0 && distinctEquipmentItems.Count() > 0)
                adj = Math.Round(cycle.ReconcileAdj / bw , 4);

            foreach (var equipmentID in distinctEquipmentItems)
            {
                var globalViewDevice = _deviceService.GetDevice(equipmentID);
                if (globalViewDevice == null) continue;

                var coFreedomDevice = coFreedomDevices.FirstOrDefault(x => x.EquipmentID == globalViewDevice.EquipmentID);

                var costByDeviceModel = new CostByDeviceModel();
                costByDeviceModel.SerialNumber = globalViewDevice.SerialNumber;
                costByDeviceModel.Model = globalViewDevice.Model;
                costByDeviceModel.DeviceType = coFreedomDevice?.ModelCategory;
                costByDeviceModel.Schedule = coFreedomDevice?.ScheduleNumber;
                costByDeviceModel.Location = coFreedomDevice?.Location;
                costByDeviceModel.User = coFreedomDevice?.AssetUser;
                costByDeviceModel.CostCenter = coFreedomDevice?.CostCenter;
                costByDeviceModel.BWVolume = GetBWVolume(items, equipmentID);
                costByDeviceModel.ColorVolume = GetColorVolume(items, equipmentID);
                if (costByDeviceModel.BWVolume > 0)
                {
                    costByDeviceModel.BWCopies = GetBwCopies(items, equipmentID, costByDeviceModel.BWVolume, contractedVolume["B/W Copies"], adj);
                    costByDeviceModel.BWPrints = GetBwPrints(items, equipmentID, costByDeviceModel.BWVolume, contractedVolume["B/W Laser Prints"], adj);
                }
                if (costByDeviceModel.ColorVolume > 0)
                {
                    costByDeviceModel.ColorPrints = GetColorPrints(items, equipmentID, costByDeviceModel.ColorVolume, contractedVolume["Color Laser Prints"], 0);
                    costByDeviceModel.ColorCopies = GetColorCopies(items, equipmentID, costByDeviceModel.ColorVolume, contractedVolume["Color Copier (Color Pages)"], 0);
                    
                }
                costByDevices.Add(costByDeviceModel);
            }

            return costByDevices.OrderBy(o=> o.DeviceType).ToList();
        }

        private static decimal GetColorCopies(IList<ViewMonthlyDeviceCosts> items, int equipmentNumber, decimal volume, DeviceCostsModel contractedcost, decimal adj)
        {
            if (volume > 0)
            {
                DeviceSummaryModel model = new DeviceSummaryModel();
                var totalvolume = items.Where(x => x.ContractMeterGroup.EqualsIgnore("Color Copier (Color Pages)")).Sum(x => x.DifferenceCopies);
                var billedpct = Math.Round((volume / totalvolume), 4);
                var billedAmt = contractedcost.Cost * billedpct;
                var overageAmt = contractedcost.Overage * billedpct;
                model.EquipmentID = equipmentNumber;
                model.MeterGroup = "Color Copier (Color Pages)";
                model.Cost = billedAmt;
                model.Overage = overageAmt;
                var results = items
                   .Where(x => x.EquipmentID == equipmentNumber)
                   .Where(x => x.ContractMeterGroup.EqualsIgnore("Color Copier (Color Pages)")).FirstOrDefault();
                if (results != null)
                {
                    results.BilledAmount = billedAmt + overageAmt;
                    
                    return results.BilledAmount + adj;
                }

            }
            return 0.00M;
            /*
            return items
                .Where(x => x.EquipmentID == equipmentNumber)
                .Where(x => x.ContractMeterGroup.EqualsIgnore("Color Copier (Color Pages)"))
                .Sum(x => x.BilledAmount);
                */
        }

        private static decimal GetColorPrints(IList<ViewMonthlyDeviceCosts> items, int equipmentNumber, decimal volume, DeviceCostsModel contractedcost, decimal adj)
        {
            if (volume > 0)
            {
                DeviceSummaryModel model = new DeviceSummaryModel();
                var totalvolume = items.Where(x => x.ContractMeterGroup.EqualsIgnore("Color Laser Prints")).Sum(x => x.DifferenceCopies);
                var billedpct = Math.Round((volume / totalvolume), 4);
                var billedAmt = contractedcost.Cost * billedpct;
                var overageAmt = contractedcost.Overage * billedpct;
                model.EquipmentID = equipmentNumber;
                model.MeterGroup = "Color Laser Prints";
                model.Cost = billedAmt;
                var results = items
                   .Where(x => x.EquipmentID == equipmentNumber)
                   .Where(x => x.ContractMeterGroup.EqualsIgnore("Color Laser Prints")).FirstOrDefault();
                if (results != null)
                {
                    results.BilledAmount = billedAmt + overageAmt;
                    return results.BilledAmount + adj;
                }

            }
            return 0.00M;
            /*
            return items
                .Where(x => x.EquipmentID == equipmentNumber)
                .Where(x => x.ContractMeterGroup.EqualsIgnore("Color Laser Prints") )
                .Sum(x => x.BilledAmount);
                */
        }

        private static decimal GetBwPrints(IList<ViewMonthlyDeviceCosts> items, int equipmentNumber, decimal volume, DeviceCostsModel contractedcost,decimal adj )
        {
            
            if (volume > 0  )
            {
                DeviceSummaryModel model = new DeviceSummaryModel();
                var totalvolume = items.Where(x => x.ContractMeterGroup.EqualsIgnore("B/W Laser Prints")).Sum(x => x.DifferenceCopies);
                var billedpct = Math.Round((volume / totalvolume), 4);
                var billedAmt = contractedcost.Cost * billedpct;
                var overageAmt = contractedcost.Overage * billedpct;
                model.EquipmentID = equipmentNumber;
                model.MeterGroup = "B/W Laser Prints";
                model.Cost = billedAmt;
                var results = items
                   .Where(x => x.EquipmentID == equipmentNumber)
                   .Where(x => x.ContractMeterGroup.EqualsIgnore("B/W Laser Prints")).FirstOrDefault();
                if (results != null)
                {
                    results.BilledAmount = billedAmt + overageAmt;
                    return results.BilledAmount + adj;
                }
                  
            }
            return 0.00M;

            /*
            return items
                .Where(x => x.EquipmentID == equipmentNumber)
                .Where(x => x.ContractMeterGroup.EqualsIgnore("B/W Laser Prints"))
                .Sum(x => x.BilledAmount);
                */
        }

        private static decimal GetBwCopies(IList<ViewMonthlyDeviceCosts> items, int equipmentNumber,decimal volume, DeviceCostsModel contractedcost,decimal adj )
        {
            
            if (volume > 0)
            {
                var totalvolume = items.Where(x => x.ContractMeterGroup.EqualsIgnore("B/W Copies")).Sum(x => x.DifferenceCopies);
                var billedpct = Math.Round((volume / totalvolume), 4);
                var billedAmt =  contractedcost.Cost * billedpct;
                var overageAmt = contractedcost.Overage * billedpct;
                var results = items
                  .Where(x => x.EquipmentID == equipmentNumber)
                  .Where(x => x.ContractMeterGroup.EqualsIgnore("B/W Copies")).FirstOrDefault();
                if (results != null)
                {
                    results.BilledAmount = billedAmt + overageAmt;
                    return results.BilledAmount + adj;
                }
            }
            return 0.00M;
            /*
            return items
                .Where(x => x.EquipmentID == equipmentNumber)
                .Where(x => x.ContractMeterGroup.EqualsIgnore("B/W Copies"))
                .Sum(x => x.BilledAmount);
            */
        }

        private static decimal GetBWVolume(IList<ViewMonthlyDeviceCosts> items, int equipmentNumber)
        {
            return items
                .Where(x => x.EquipmentID == equipmentNumber)
                .Where(x => x.ContractMeterGroup.EqualsIgnore("B/W Copies") || x.ContractMeterGroup.EqualsIgnore("B/W Laser Prints"))
                .Sum(x => x.DifferenceCopies);
        }
        private static decimal GetColorVolume(IList<ViewMonthlyDeviceCosts> items, int equipmentNumber)
        {
            return items
                .Where(x => x.EquipmentID == equipmentNumber)
                .Where(x => x.ContractMeterGroup.EqualsIgnore("Color Copier (Color Pages)") || x.ContractMeterGroup.EqualsIgnore("Color Laser Prints"))
                .Sum(x => x.DifferenceCopies);
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