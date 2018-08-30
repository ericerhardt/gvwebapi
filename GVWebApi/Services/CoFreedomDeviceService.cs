using System;
using System.Collections.Generic;
using System.Linq;
using GV.Domain;
using GV.Domain.Entities;
using GVWebapi.Models.Schedules;
using GVWebapi.RemoteData;

namespace GVWebapi.Services
{
    public interface ICoFreedomDeviceService
    {
        void LoadCoFreedomDevices(long scheduleId);
        vw_admin_EquipmentList_MeterGroup GetCoFreedomDevice(long EquipmentId);
        Dictionary<string, int> GetDeviceCount(long customerId);
        int GetScheduleDeviceCount(string scheduleNumber);
        EditScheduleDeviceTopModel GetTabCounts(long scheduleId);
        IList<vw_admin_EquipmentList_MeterGroup> GetCoFreedomDevices(string scheduleName, long customerId);
        IList<vw_admin_EquipmentList_MeterGroup> GetCoFreedomDevicesNoSchedule(long customerId);
        IList<vw_admin_EquipmentList_MeterGroup> GetCoFreedomDevices(long customerId);
    }

    public class CoFreedomDeviceService : ICoFreedomDeviceService
    {
        private readonly IRepository _repository;

        public CoFreedomDeviceService(IRepository repository)
        {
            _repository = repository;
        }

        public void LoadCoFreedomDevices(long scheduleId)
        {
            var schedule = GetSchedule(scheduleId);
            if (schedule == null) return;

            var coFreedomDevices = GetCoFreedomDevices(schedule.CustomerId, schedule.Name);

            AddUpdateDevices(coFreedomDevices, schedule);
            AddUnAllocatedDevices(coFreedomDevices, schedule);
            SetRemovedDevices(coFreedomDevices, schedule);
        }
        public int GetScheduleDeviceCount(string scheduleNumber)
        {
            try
            {
                using (var freedomEntities = new CoFreedomEntities())
                {
                   var count = freedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.ScheduleNumber == scheduleNumber && x.NumberOfContractsActive != null).Count();
                    return count;
                }
            }
            catch
            {
                return 0;
            }
          
        }
        public Dictionary<string, int> GetDeviceCount(long customerId)
        {
            try
            {


                using (var freedomEntities = new CoFreedomEntities())
                {
                    return freedomEntities
                        .vw_admin_EquipmentList_MeterGroup
                        .Where(x => x.CustomerID == customerId && x.ScheduleNumber != null)
                        .GroupBy(x => x.ScheduleNumber)
                        .Select(x => new
                        {
                            Name = x.Key,
                            Count = x.Count()
                        })
                        .ToDictionary(x => x.Name, x => x.Count);
                }
            } catch 
            {
                return null;
            }
        }

        public EditScheduleDeviceTopModel GetTabCounts(long scheduleId)
        {
            var schedule = GetSchedule(scheduleId);
            var model = new EditScheduleDeviceTopModel();
            var coFreedomDevices = GetCoFreedomDevices(schedule.CustomerId, schedule.Name);
            model.ActiveCount = coFreedomDevices.Count(x => x.ScheduleNumber != null);
            model.UnAllocatedCount = coFreedomDevices.Count(x => x.ScheduleNumber == null);
            model.RemovedCount = schedule.Devices.Count(x => x.RemovedStatus == RemovedStatusEnum.Removed || x.RemovedStatus == RemovedStatusEnum.FormatterReplaced);
            return model;
        }

        public IList<vw_admin_EquipmentList_MeterGroup> GetCoFreedomDevices(string scheduleName, long customerId)
        {
            using (var freedomEntities = new CoFreedomEntities())
            {
                return freedomEntities
                    .vw_admin_EquipmentList_MeterGroup
                    .Where(x => x.CustomerID == customerId)
                    .Where(x => x.ScheduleNumber.ToLower().Trim() == scheduleName.ToLower().Trim() && x.NumberOfContractsActive != null)
                    .ToList();
            }
        }
        public vw_admin_EquipmentList_MeterGroup GetCoFreedomDevice(long EquipmentId)
        {
            using (var freedomEntities = new CoFreedomEntities())
            {
                return freedomEntities
                    .vw_admin_EquipmentList_MeterGroup
                    .Where(x => x.EquipmentID == EquipmentId)
                  
                    .FirstOrDefault();
            }
        }
        public IList<vw_admin_EquipmentList_MeterGroup> GetCoFreedomDevicesNoSchedule(long customerId)
        {
            using (var freedomEntities = new CoFreedomEntities())
            {
                return freedomEntities
                    .vw_admin_EquipmentList_MeterGroup
                    .Where(x => x.CustomerID == customerId)
                    .Where(x => x.ScheduleNumber.ToLower().Trim().Length == 0)
                    .ToList();
            }
        }

        public IList<vw_admin_EquipmentList_MeterGroup> GetCoFreedomDevices(long customerId)
        {
            using (var freedomEntities = new CoFreedomEntities())
            {
                return freedomEntities
                    .vw_admin_EquipmentList_MeterGroup
                    .Where(x => x.CustomerID == customerId)
                    .ToList();
            }
        }

        private static void SetRemovedDevices(IList<CoFreedomDeviceModel> coFreedomDevices, SchedulesEntity schedule)
        {
            foreach (var scheduleDevice in schedule.Devices)
            {
                if (scheduleDevice.RemovedStatus == RemovedStatusEnum.Removed) continue;
                if (scheduleDevice.RemovedStatus == RemovedStatusEnum.FormatterReplaced) continue;
                var existsInCoFreedom = coFreedomDevices.FirstOrDefault(x => x.EquipmentId == scheduleDevice.EquipmentId);
                if (existsInCoFreedom != null) continue;
                scheduleDevice.RemovedStatus = RemovedStatusEnum.SetForRemoval;
                scheduleDevice.ModifiedDateTime = DateTimeOffset.Now;
            }
        }

        private SchedulesEntity GetSchedule(long scheduleId)
        {
            return _repository.Get<SchedulesEntity>(scheduleId);
        }

        private void AddUnAllocatedDevices(IList<CoFreedomDeviceModel> coFreedomDevices, SchedulesEntity schedule)
        {
            var unAllocatedDevices = _repository.Find<DevicesEntity>()
                .Where(x => x.CustomerId == schedule.CustomerId)
                .Where(x => x.Schedule == null)
                .ToList();

            foreach (var coFreedomDeviceModel in coFreedomDevices.Where(x => x.ScheduleNumber == null))
            {
                var existingGlobalViewDevice = unAllocatedDevices.FirstOrDefault(x => x.EquipmentId == coFreedomDeviceModel.EquipmentId);
                if (existingGlobalViewDevice == null)
                {
                    //add unallocated device
                    var newDevice = new DevicesEntity
                    {
                        CustomerId = schedule.CustomerId,
                        EquipmentId = coFreedomDeviceModel.EquipmentId,
                        EquipmentNumber = coFreedomDeviceModel.EquipmentNumber,
                        SerialNumber = coFreedomDeviceModel.SerialNumber,
                        Model = coFreedomDeviceModel.Model,
                        CreatedDateTime = DateTimeOffset.Now
                    };
                    _repository.Add(newDevice);
                }
                else
                {
                    //update unallocated devices
                    if (existingGlobalViewDevice.ChangeKey != coFreedomDeviceModel.ChangeKey)
                    {
                        existingGlobalViewDevice.EquipmentNumber = coFreedomDeviceModel.EquipmentNumber;
                        existingGlobalViewDevice.SerialNumber = coFreedomDeviceModel.SerialNumber;
                        existingGlobalViewDevice.Model = coFreedomDeviceModel.Model;
                        existingGlobalViewDevice.ModifiedDateTime = DateTimeOffset.Now;
                    }
                }
            }
        }

        private static void AddUpdateDevices(IList<CoFreedomDeviceModel> coFreedomDevices, SchedulesEntity schedule)
        {
            foreach (var deviceModel in coFreedomDevices)
            {
                if (deviceModel.ScheduleNumber == null) continue;
                var existingDevice = schedule.Devices
                    .Where(x => x.Schedule.Name == schedule.Name)
                    .FirstOrDefault(x => x.EquipmentId == deviceModel.EquipmentId);

                if (existingDevice == null)
                {
                    //add new device
                    var newDevice = new DevicesEntity
                    {
                        CustomerId = schedule.CustomerId,
                        EquipmentId = deviceModel.EquipmentId,
                        EquipmentNumber = deviceModel.EquipmentNumber,
                        SerialNumber = deviceModel.SerialNumber,
                        Model = deviceModel.Model,
                        CreatedDateTime = DateTimeOffset.Now
                    };
                    schedule.AddDevice(newDevice);
                }
                else
                {
                    if (existingDevice.ChangeKey != deviceModel.ChangeKey)
                    {
                        existingDevice.EquipmentNumber = deviceModel.EquipmentNumber;
                        existingDevice.SerialNumber = deviceModel.SerialNumber;
                        existingDevice.Model = deviceModel.Model;
                        existingDevice.ModifiedDateTime = DateTimeOffset.Now;
                    }
                }
            }
        }

        private static IList<CoFreedomDeviceModel> GetCoFreedomDevices(long customerId, string scheduleName)
        {
            using (var freedomEntities = new CoFreedomEntities())
            {
                return freedomEntities
                    .vw_admin_EquipmentList_MeterGroup
                    .Where(x => x.CustomerID == customerId)
                    .Where(x => x.ScheduleNumber.ToLower().Trim() == scheduleName.ToLower().Trim() || x.ScheduleNumber.Trim().Length == 0)
                    .Select(x => new CoFreedomDeviceModel
                    {
                        ScheduleNumber = x.ScheduleNumber.Trim().Length == 0 ? null : x.ScheduleNumber,
                        EquipmentId = x.EquipmentID,
                        EquipmentNumber = x.EquipmentNumber,
                        SerialNumber = x.SerialNumber,
                        Model = x.Model,
                        IsActive = x.Active
                    }).ToList();
            }
        }
    }

    public class CoFreedomDeviceModel
    {
        public string ScheduleNumber { get; set; }
        public int EquipmentId { get; set; }
        public string EquipmentNumber { get; set; }
        public string SerialNumber { get; set; }
        public string Model { get; set; }
        public bool IsActive { get; set; }
        public string ChangeKey => $"{EquipmentId}/{EquipmentNumber.Trim().ToLower()}/{SerialNumber.Trim().ToLower()}{Model.Trim().ToLower()}";
    }
}