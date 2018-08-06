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
        IList<DeviceModel> GetRemovedDevices(long scheduleId);
        void DeleteDevice(long deviceId);
        void SaveDevice(DeviceSaveModel model);
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

        public DeviceService(IRepository repository, ICoFreedomDeviceService coFreedomDeviceService, ICoFreedomRepository coFreedomRepository)
        {
            _repository = repository;
            _coFreedomDeviceService = coFreedomDeviceService;
            _coFreedomRepository = coFreedomRepository;
        }

        public IList<DeviceModel> GetActiveDevices(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);

            var globalViewEntities = _repository.Find<DevicesEntity>()
                .Where(x => x.Schedule.ScheduleId == scheduleId)
                .Where(x => x.RemovedStatus == null || x.RemovedStatus == RemovedStatusEnum.SetForRemoval)
                .ToList();
            var scheduldeName = schedule.Name.Insert(9, "-");

            var coFreedomDevices = _coFreedomDeviceService.GetCoFreedomDevices(scheduldeName, schedule.CustomerId);

            return MergeGlobalViewAndCoFreedom(globalViewEntities, coFreedomDevices);

        }

        private static List<DeviceModel> MergeGlobalViewAndCoFreedom(List<DevicesEntity> globalViewEntities, IList<vw_admin_EquipmentList_MeterGroup> coFreedomDevices)
        {
            var devices = new List<DeviceModel>();
            foreach (var globalViewEntity in globalViewEntities)
            {
                var existingCoFreedomDevice = coFreedomDevices.FirstOrDefault(x => x.EquipmentID == globalViewEntity.EquipmentId);
                if (existingCoFreedomDevice == null) continue;
                devices.Add(DeviceModel.For(globalViewEntity, existingCoFreedomDevice));
            }

            return devices;
        }

        public IList<DeviceModel> GetUnallocatedDevices(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);

            var globalViewEntities = _repository.Find<DevicesEntity>()
                .Where(x => x.Schedule == null)
                .Where(x => x.CustomerId == schedule.CustomerId)
                .ToList();

            var coFreedomDevices = _coFreedomDeviceService.GetCoFreedomDevicesNoSchedule(schedule.CustomerId);

            return MergeGlobalViewAndCoFreedom(globalViewEntities, coFreedomDevices);
        }

        public IList<DeviceModel> GetRemovedDevices(long scheduleId)
        {
            var schedule = _repository.Get<SchedulesEntity>(scheduleId);

            var globalViewEntities = _repository.Find<DevicesEntity>()
                .Where(x => x.RemovedStatus == RemovedStatusEnum.Removed)
                .Where(x => x.CustomerId == schedule.CustomerId)
                .ToList();

            var coFreedomDevices = _coFreedomDeviceService.GetCoFreedomDevices(schedule.CustomerId);

            return MergeGlobalViewAndCoFreedom(globalViewEntities, coFreedomDevices);
        }

        public void DeleteDevice(long deviceId)
        {
            var device = _repository.Get<DevicesEntity>(deviceId);
            if (device == null) return;
            var customProperty = _coFreedomRepository.Find<ScEquipmentCustomPropertiesEntity>()
                .Where(x => x.Equipment.EquipmentId == device.EquipmentId)
                .FirstOrDefault(x => x.ShAttributeId == 2015);
            if (customProperty == null) return;
            customProperty.TextVal = string.Empty;
        }

        public void SaveDevice(DeviceSaveModel model)
        {
            var device = _repository.Get<DevicesEntity>(model.DeviceId);
            if (device == null) return;

            device.MonthlyCost = model.MonthlyCost;

            var scEquipmentEntity = _coFreedomRepository.Get<ScEquipmentEntity>(device.EquipmentId);

            //set schedule -- custom prop 2015
            var scheduleProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2015);
            if (scheduleProperty != null)
                scheduleProperty.TextVal = device.Schedule.Name;

            scEquipmentEntity.Location = _coFreedomRepository.Load<ArCustomersEntity>(model.LocationId);

            var userProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2014);
            if (userProperty != null)
                userProperty.TextVal = model.User;

            var exhibitProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2006);
            if (exhibitProperty != null)
                exhibitProperty.TextVal = $"Exhibit {model.Exhibit}";

            var costCenterProperty = scEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2001);
            if (costCenterProperty != null)
                costCenterProperty.TextVal = model.CostCenter;
        }

        public void AddDevicesToSchedule(SetScheduleSaveModel model)
        {
            var schedule = _repository.Get<SchedulesEntity>(model.ScheduleId);

            foreach (var deviceId in model.DeviceIds)
            {
                var deviceEntity = _repository.Get<DevicesEntity>(deviceId);
                if (deviceEntity == null) continue;
                var coFreedomEquipmentEntity = _coFreedomRepository.Get<ScEquipmentEntity>(deviceEntity.EquipmentId);
                var scheduleCustomProperty = coFreedomEquipmentEntity.CustomProperties.FirstOrDefault(x => x.ShAttributeId == 2015);
                if (scheduleCustomProperty != null)
                    scheduleCustomProperty.TextVal = schedule.Name;
            }

            _coFreedomDeviceService.LoadCoFreedomDevices(model.ScheduleId);
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
    }

    public enum DeviceStatusEnum
    {
        Removed,
        Active,
        InActive
    }
}