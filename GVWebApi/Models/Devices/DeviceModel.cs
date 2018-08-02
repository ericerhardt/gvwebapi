using System;
using GV.Domain.Entities;
using GVWebapi.RemoteData;
using GVWebapi.Services;

namespace GVWebapi.Models.Devices
{
    public class DeviceModel
    {
        public static DeviceModel For(DevicesEntity deviceEntity, vw_admin_EquipmentList_MeterGroup coFreedomDevice)
        {
            var model = new DeviceModel();
            model.DeviceId = deviceEntity.DeviceId;
            model.EquipmentId = deviceEntity.EquipmentId;
            model.EquipmentNumber = deviceEntity.EquipmentNumber;
            model.SerialNumber = deviceEntity.SerialNumber;
            model.Model = deviceEntity.Model;
            model.ScheduleNumber = deviceEntity.Schedule?.Name;
            model.Exhibit = GetExhibit(coFreedomDevice);
            model.Location = coFreedomDevice.LocName;
            model.User = coFreedomDevice.AssetUser;
            model.CostCenter = coFreedomDevice.CostCenter;
            model.Status = GetDeviceStatus(coFreedomDevice, deviceEntity);
            model.MonthlyCost = deviceEntity.MonthlyCost;
            model.Disposition = deviceEntity.Disposition;
            model.RemovedDateTime = deviceEntity.RemovalDateTime?.LocalDateTime;
            model.DeviceType = coFreedomDevice.ModelCategory;

            return model;
        }

        private static DeviceStatusEnum GetDeviceStatus(vw_admin_EquipmentList_MeterGroup coFreedomDevice, DevicesEntity deviceEntity)
        {
            if (deviceEntity.RemovedStatus.HasValue && deviceEntity.RemovedStatus.Value == RemovedStatusEnum.SetForRemoval)
            {
                return DeviceStatusEnum.Removed;
            }

            switch (coFreedomDevice.Active)
            {
                case true:
                    return DeviceStatusEnum.Active;
                case false:
                    return DeviceStatusEnum.InActive;
                default:
                    return DeviceStatusEnum.Removed;
            }
        }

        private static string GetExhibit(vw_admin_EquipmentList_MeterGroup coFreedomDevice)
        {
            switch (coFreedomDevice.OwnershipType.Trim().ToLower())
            {
                case "":
                case "client owned":
                    return "B";
                case "exhibit a":
                    return "A";
                default:
                    return string.Empty;
            }
        }

        private DeviceModel()
        {
        }

        public long DeviceId { get; set; }
        public int EquipmentId { get; set; }
        public string EquipmentNumber { get; set; }
        public string SerialNumber { get; set; }
        public string Model { get; set; }
        public string ScheduleNumber { get; set; }
        public string Exhibit { get; set; }
        public string Location { get; set; }
        public string User { get; set; }
        public string CostCenter { get; set; }
        public decimal MonthlyCost { get; set; }
        public DeviceStatusEnum Status { get; set; }
        public string Disposition { get; set; }
        public DateTime? RemovedDateTime { get; set; }
        public string DeviceType {get; set; }
        public decimal TaxRate { get; set; }
        public decimal CalculatedTax => MonthlyCost * TaxRate;
        public decimal UnitTotal => CalculatedTax + MonthlyCost;
    }
}