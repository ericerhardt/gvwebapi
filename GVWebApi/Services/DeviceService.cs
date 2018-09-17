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
    public interface IDeviceService
    {
        IList<DeviceModel> GetActiveDevices(long scheduleId);
        IList<DeviceModel> GetUnallocatedDevices(long scheduleId);

        DeviceModel GetDeviceByID(long deviceId);
        void DeleteDevice(long deviceId);
        void SaveDevice(DeviceSaveModel model);
        decimal DeviceTotalCost(long scheduleId);
        void AddDevicesToSchedule(SetScheduleSaveModel model);
        void ConfirmDeviceRemove(DeviceRemoveModel model);
        void ConfirmFormatterReplacement(FormatterReplacedModel model);
        IList<DeviceSearchModel> GetDevicesToSearch(long scheduleId);
        void AddReplacementDevice(DeviceReplacementSaveModel model);
        DevicesEntity GetDevice(string equipmentNumber);
       
}

    public class DeviceService : IDeviceService
    {
        private readonly IRepository _repository;
        private readonly ICoFreedomDeviceService _coFreedomDeviceService;
        private readonly ICoFreedomRepository _coFreedomRepository;
        private readonly ILocationsService _locationsService;
        public DeviceService(IRepository repository, ICoFreedomDeviceService coFreedomDeviceService, ICoFreedomRepository coFreedomRepository, ILocationsService locationsService)
        {
            _repository = repository;
            _coFreedomDeviceService = coFreedomDeviceService;
            _coFreedomRepository = coFreedomRepository;
            _locationsService = locationsService;
        }

        public IList<DeviceModel> GetActiveDevices(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);
            var devices = new List<DeviceModel>();
            var coFreedomDevices = _coFreedomDeviceService.GetCoFreedomDevices(schedule.Name, schedule.CustomerId);
            foreach(var device in coFreedomDevices)
            {
                var taxrate = _locationsService.GetTaxRate(device.LocName);
                devices.Add(DeviceModel.For(taxrate,device));
            }

            return devices;

        }

        //private static List<DeviceModel> MergeGlobalViewAndCoFreedom(List<DevicesEntity> globalViewEntities, IList<vw_admin_EquipmentList_MeterGroup> coFreedomDevices)
        //{
        //    var devices = new List<DeviceModel>();
        //    foreach (var globalViewEntity in globalViewEntities)
        //    {
        //        var existingCoFreedomDevice = coFreedomDevices.FirstOrDefault(x => x.EquipmentID == globalViewEntity.EquipmentId);
        //        if (existingCoFreedomDevice == null) continue;
        //        devices.Add(DeviceModel.For(globalViewEntity, existingCoFreedomDevice));
        //    }

        //    return devices;
        //}
        //private static  DeviceModel  MergeGlobalViewAndCoFreedomDevice(DevicesEntity globalViewEntity, vw_admin_EquipmentList_MeterGroup coFreedomDevice)
        //{ 
        //    var device = DeviceModel.For(globalViewEntity, coFreedomDevice);
        //    return device;
        //}


        public IList<DeviceModel> GetUnallocatedDevices(long customerId)
        {
            
            var globalViewEntities = _repository.Find<vw_admin_EquipmentList_MeterGroup>()
                .Where(x => x.ScheduleNumber == null)
                .Where(x => x.CustomerID == customerId)
                .ToList();
            var devices = new List<DeviceModel>();
            foreach (var device in globalViewEntities)
            {
                var taxrate = _locationsService.GetTaxRate(device.LocName);
                devices.Add(DeviceModel.For(taxrate,device));
            }

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

            var devices = _coFreedomDeviceService.GetCoFreedomDevices(schedule.Name, schedule.CustomerId);
                
            var totals = 0.00M;
            foreach(var device in devices)
            {
                totals = totals + decimal.Parse(device.MonthlyCost);
            }
           
            return totals;
        }
        public void SaveDevice(DeviceSaveModel model)
        {

            var schedulesEntity = _repository.Get<SchedulesEntity>(model.ScheduleId);
            var scEquipmentEntity = _coFreedomRepository.Get<ScEquipmentEntity>(model.DeviceId);

            //set schedule -- custom prop 2015
            var scheduleProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2015);
            if (scheduleProperty != null)
                scheduleProperty.TextVal = schedulesEntity.Name;

            scEquipmentEntity.Location = _coFreedomRepository.Load<ArCustomersEntity>(model.LocationId);

            var userProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2014);
            if (userProperty != null)
                userProperty.TextVal = model.User;

            var exhibitProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2006);
            if (exhibitProperty != null)
                exhibitProperty.TextVal =  model.Exhibit;

            var costCenterProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2001);
            if (costCenterProperty != null)
                costCenterProperty.TextVal = model.CostCenter;

            var monthlyCostProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2021);
            if (monthlyCostProperty != null)
                monthlyCostProperty.TextVal = model.MonthlyCost.ToString();
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

          //  _coFreedomDeviceService.LoadCoFreedomDevices(model.ScheduleId);
        }

        public void ConfirmDeviceRemove(DeviceRemoveModel model)
        {
            var device = _repository.Get<DevicesEntity>(model.DeviceId);
            device.ModifiedDateTime = DateTimeOffset.Now;
            device.RemovedStatus = RemovedStatusEnum.Removed;
            device.RemovalDateTime = new DateTimeOffset(model.RemovalDate.Date);
            device.Disposition = "Device Removed";
        }

        public void ConfirmFormatterReplacement(FormatterReplacedModel model)
        {
            var device = _repository.Get<DevicesEntity>(model.DeviceId);
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
            var oldDevice = _repository.Get<DevicesEntity>(model.OldDeviceId);
            oldDevice.RemovalDateTime = DateTimeOffset.Now;
            oldDevice.ModifiedDateTime = DateTimeOffset.Now;
            oldDevice.RemovedStatus = RemovedStatusEnum.Removed;
            oldDevice.Disposition = $"Device Replaced by SN#{model.NewSerialNumber}";

            var replacementEntity = new AssetReplacementEntity();
            replacementEntity.CustomerId = Convert.ToInt32(oldDevice.CustomerId);
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

        public DevicesEntity GetDevice(string equipmentNumber)
        {
            return _repository.Find<DevicesEntity>()
                .FirstOrDefault(x => x.EquipmentNumber.ToLower() == equipmentNumber.ToLower());
        }
        public DeviceModel GetDeviceByID(long EquipmentID)
        {
            var eaDevice = _coFreedomDeviceService.GetCoFreedomDevice(EquipmentID);
            var taxrate = _locationsService.GetTaxRate(eaDevice.LocName);
            var Device = DeviceModel.For(taxrate,eaDevice);
            return Device; 

        }

    }

    public enum DeviceStatusEnum
    {
        Removed,
        Active,
        InActive
    }
}