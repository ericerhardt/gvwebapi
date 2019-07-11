using System;
using System.Collections.Generic;
using System.Linq;
using GV.CoFreedomDomain;
using GV.CoFreedomDomain.Entities;
using GV.Domain;
using GV.Domain.Entities;
using GVWebapi.Models.Devices;
using GVWebapi.RemoteData;

namespace GVWebapi.Services
{
    public interface IScheduleDevicesService
    {
        IList<ScheduleDevicesModel> GetActiveDevices(long scheduleId);
        IList<vw_admin_EquipmentList_MeterGroup> GetUnallocatedDevices(long customerId);

        ScheduleDevicesModel GetDeviceByID(long deviceId);
        void DeleteDevice(long deviceId);
        void SaveDevice(DeviceSaveModel model);
        decimal DeviceTotalCost(long scheduleId);
        void AddDevicesToSchedule(SetScheduleSaveModel model);
        void ConfirmDeviceRemove(DeviceRemoveModel model);
        void ConfirmFormatterReplacement(FormatterReplacedModel model);
        IList<DeviceSearchModel> GetDevicesToSearch(long scheduleId);
        void AddReplacementDevice(DeviceReplacementSaveModel model);
        vw_admin_EquipmentList_MeterGroup GetDevice(int equipmentID);
        IList<ScheduleDevicesModel> GetScheduleDevices(long CyclePeriodId, long CustomerId);
}

    public class ScheduleDevicesService : IScheduleDevicesService
    {
        private readonly IRepository _repository;
        private readonly ICoFreedomDeviceService _coFreedomDeviceService;
        private readonly ICoFreedomRepository _coFreedomRepository;
        private readonly ILocationsService _locationsService;
        public ScheduleDevicesService(IRepository repository, ICoFreedomDeviceService coFreedomDeviceService, ICoFreedomRepository coFreedomRepository, ILocationsService locationsService)
        {
            _repository = repository;
            _coFreedomDeviceService = coFreedomDeviceService;
            _coFreedomRepository = coFreedomRepository;
            _locationsService = locationsService;
        }
        public IList<ScheduleDevicesModel> GetScheduleDevices(long CyclePeriodId, long CustomerId)
        {
           var schedule = _repository.Find<CyclePeriodSchedulesEntity>()
                                      .Where(x=> x.CyclePeriod.CyclePeriodId == CyclePeriodId)
                                      .Select(g => g.Schedule.ScheduleId)
                                      .ToList();

            var devices = new List<ScheduleDevicesModel>();
            var serivceDevices = _repository.Find<ScheduleDevicesEntity>()
                                            .Where(x => schedule.Contains(x.Schedule.ScheduleId)).ToList();
            foreach (var device in serivceDevices)
            {
                var taxrate = _locationsService.GetTaxRate(device.Location);
                var instance = InstanceMultipler2(CyclePeriodId);

                devices.Add(ScheduleDevicesModel.For(taxrate, instance, device));
            }

            return devices;

        }
        public IList<ScheduleDevicesModel> GetActiveDevices(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);
            
            var devices = new List<ScheduleDevicesModel>();
            var scheduleDevices = _repository.Find<ScheduleDevicesEntity>()
                                              .Where(x => x.CustomerID == schedule.CustomerId)
                                              .Where(x => x.ScheduleNumber == schedule.Name);
            foreach(var device in scheduleDevices)
            {
                var taxrate = _locationsService.GetTaxRate(device.Location);
                var instance = InstanceMultipler(scheduleId);
               
                devices.Add(ScheduleDevicesModel.For(taxrate,instance,device));
            }

            return devices;

        }
        private decimal InstanceMultipler2(long cyclePeriodId)
        {
            using (var context = new GlobalViewEntities())
            {
                var cycleperiod = context.CyclePeriodSchedules.Where(x => x.CyclePeriodId == cyclePeriodId).FirstOrDefault();
                return cycleperiod.InstancesInvoiced;
            }

        }
        private decimal InstanceMultipler(long scheduleId)
        {
            using(var context = new GlobalViewEntities())
            {
                var cycleperiod = context.CyclePeriodSchedules.Where(x => x.ScheduleId == scheduleId).OrderByDescending(x => x.CyclePeriodScheduleId).FirstOrDefault();
                return cycleperiod.InstancesInvoiced;
            }
           
        }
        //private static List<ScheduleDevicesModel> MergeGlobalViewAndCoFreedom(List<ScheduleDevicesEntity> globalViewEntities, IList<vw_admin_EquipmentList_MeterGroup> coFreedomDevices)
        //{
        //    var devices = new List<ScheduleDevicesModel>();
        //    foreach (var globalViewEntity in globalViewEntities)
        //    {
        //        var existingCoFreedomDevice = coFreedomDevices.FirstOrDefault(x => x.EquipmentID == globalViewEntity.EquipmentID);
        //        if (existingCoFreedomDevice == null) continue;
        //        devices.Add(ScheduleDevicesModel.For(globalViewEntity, existingCoFreedomDevice));
        //    }

        //    return devices;
        //}
        //private static  DeviceModel  MergeGlobalViewAndCoFreedomDevice(DevicesEntity globalViewEntity, vw_admin_EquipmentList_MeterGroup coFreedomDevice)
        //{ 
        //    var device = DeviceModel.For(globalViewEntity, coFreedomDevice);
        //    return device;
        //}


        public IList<vw_admin_EquipmentList_MeterGroup> GetUnallocatedDevices(long customerId)
        {
            
            var devices = _coFreedomRepository.Find<vw_admin_EquipmentList_MeterGroup>()
                .Where(x => x.ScheduleNumber == null)
                .Where(x => x.CustomerID == customerId)
                .ToList();

            return devices;
        }

       

        public void DeleteDevice(long EquipmentId)
        {
          
            var customProperty = _coFreedomRepository.Find<ScEquipmentCustomPropertiesEntity>()
                .Where(x => x.Equipment.EquipmentId ==  EquipmentId)
                .FirstOrDefault(x => x.ShAttributeId == 2015);
            if (customProperty == null) return;
            customProperty.TextVal = string.Empty;
            
        }
        public decimal DeviceTotalCost(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);

            var devices = _coFreedomDeviceService.GetScheduleDevices(schedule.ScheduleId, schedule.CustomerId);
                
            var totals = 0.00M;
            foreach(var device in devices)
            {
                totals = totals + device.MonthlyCost;
            }
           
            return totals;
        }
        //public void SaveDevice(DeviceSaveModel model)
        //{

        //    var schedulesEntity = _repository.Get<SchedulesEntity>(model.ScheduleId);
        //    var scEquipmentEntity = _coFreedomRepository.Get<ScEquipmentEntity>(model.EquipmentID);

        //    //set schedule -- custom prop 2015
        //    var scheduleProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2015);
        //    if (scheduleProperty != null)
        //        scheduleProperty.TextVal = schedulesEntity.Name;

        //    scEquipmentEntity.Location = _coFreedomRepository.Load<ArCustomersEntity>(model.LocationId);

        //    var userProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2014);
        //    if (userProperty != null)
        //        userProperty.TextVal = model.User;

        //    var exhibitProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2006);
        //    if (exhibitProperty != null)
        //        exhibitProperty.TextVal =  model.Exhibit;

        //    var costCenterProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2001);
        //    if (costCenterProperty != null)
        //        costCenterProperty.TextVal = model.CostCenter;

        //    var monthlyCostProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2021);
        //    if (monthlyCostProperty != null)
        //        monthlyCostProperty.TextVal = model.MonthlyCost.ToString();
        //}
        
        public void SaveDevice(DeviceSaveModel model)
        {

            var schedulesEntity = _repository.Get<SchedulesEntity>(model.ScheduleId);
            var Location = _coFreedomRepository.Load<ArCustomersEntity>(model.LocationId);
            var scheduleDevice = _repository.Get<ScheduleDevicesEntity>(model.EquipmentID);
            if(scheduleDevice != null)
            {
                scheduleDevice.ScheduleNumber = schedulesEntity.Name;
                scheduleDevice.Location = Location.CustomerName;
                scheduleDevice.LocationID = model.LocationId;
                scheduleDevice.AssetUser = model.User;
                scheduleDevice.OwnershipType = model.Exhibit;
                scheduleDevice.MonthlyCost = model.MonthlyCost;
                scheduleDevice.ModifiedDateTime = DateTime.Now;
            }

             
        }

        public void AddDevicesToSchedule(SetScheduleSaveModel model)
        {
            var schedule = _repository.Get<SchedulesEntity>(model.ScheduleId);

            foreach (var EquipmentId in model.DeviceIds)
            {
               
                
                var coFreedomEquipmentEntity = _coFreedomRepository.Get<ScEquipmentEntity>(EquipmentId);
                var scheduleCustomProperty = coFreedomEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2015);
                if (scheduleCustomProperty != null)
                    scheduleCustomProperty.TextVal = schedule.Name;
            }

            // _coFreedomDeviceService.LoadCoFreedomDevices(model.ScheduleId);
        }

        public void ConfirmDeviceRemove(DeviceRemoveModel model)
        {
            var device = _repository.Get<ScheduleDevicesEntity>(model.DeviceId);
            device.ModifiedDateTime = DateTimeOffset.Now;
            device.RemovedStatus = RemovedStatusEnum.Removed;
            device.RemovalDateTime = new DateTimeOffset(model.RemovalDate.Date);
            device.Disposition = "Device Removed";
        }

        public void ConfirmFormatterReplacement(FormatterReplacedModel model)
        {
            var device = _repository.Get<ScheduleDevicesEntity>(model.DeviceId);
            device.ModifiedDateTime = DateTimeOffset.Now;
            device.RemovedStatus = RemovedStatusEnum.FormatterReplaced;
            device.RemovalDateTime = new DateTimeOffset(model.ReplacementDate.Date);
            device.Disposition = "Formatter Replaced";
        }

        public IList<DeviceSearchModel> GetDevicesToSearch(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);

            return _coFreedomRepository.Find<ScEquipmentEntity>()
                .Where(x => x.CustomerId == schedule.CustomerId)
                .Select(x => new DeviceSearchModel
                {
                    EquipmentId = x.EquipmentId,
                    SerialNumber = x.SerialNumber,
                    Model = x.Model.Model,
                    Location = x.Location.CustomerName,
                    EquipmentNumber = x.EquipmentNumber,
                    Status = x.Active ? "Active" : "InActive"
                })
                .ToList();
        }

        public void AddReplacementDevice(DeviceReplacementSaveModel model)
        {
            var oldDevice = _repository.Get<ScheduleDevicesEntity>(model.OldDeviceId);
            oldDevice.RemovalDateTime = DateTimeOffset.Now;
            oldDevice.ModifiedDateTime = DateTimeOffset.Now;
            oldDevice.RemovedStatus = RemovedStatusEnum.Removed;
            oldDevice.Disposition = $"Device Replaced by SN#{model.NewSerialNumber}";

            var replacementEntity = new AssetReplacementEntity();
            replacementEntity.CustomerId = Convert.ToInt32(oldDevice.CustomerID);
            replacementEntity.Location = model.Location;
            replacementEntity.ReplacementDate = DateTime.Now;
            replacementEntity.OldSerialNumber = oldDevice.SerialNumber;
            replacementEntity.OldModel = oldDevice.Model;
            replacementEntity.NewSerialNumber = model.NewSerialNumber;
            replacementEntity.NewModel = model.NewModel;
            replacementEntity.ReplacementValue = model.ReplacementValue;
            replacementEntity.ScheduleNumber = model.ScheduleNumber;
            _repository.Add(replacementEntity);
        }

        public vw_admin_EquipmentList_MeterGroup GetDevice(int equipmentID)
        {
            return  _coFreedomDeviceService.GetCoFreedomDevice(equipmentID);
               
        }
        public ScheduleDevicesModel GetDeviceByID(long scheduleDeviceID)
        {
            // var eaDevice = _coFreedomDeviceService.GetCoFreedomDevice(EquipmentID);
            var scheduleDevice = _repository.Get<ScheduleDevicesEntity>(scheduleDeviceID);
            var taxrate = _locationsService.GetTaxRate(scheduleDevice.Location);
            var Device = ScheduleDevicesModel.For(taxrate,1, scheduleDevice);
            return Device; 

        }

    }

    public enum ScheduleDevicesStatusEnum
    {
        Removed,
        Active,
        InActive
    }
}