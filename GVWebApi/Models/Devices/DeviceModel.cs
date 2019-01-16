using System;
using GV.Domain.Entities;
using GVWebapi.RemoteData;
using GVWebapi.Services;

namespace GVWebapi.Models.Devices
{
    public class DeviceModel
    {
        public static DeviceModel For(decimal taxRate,decimal instance, vw_admin_EquipmentList_MeterGroup coFreedomDevice)
        {
            Decimal.TryParse(coFreedomDevice.MonthlyCost, out decimal MontlyCost);
            var model = new DeviceModel();
            model.EquipmentId = coFreedomDevice.EquipmentID;
            model.EquipmentNumber = coFreedomDevice.EquipmentNumber;
            model.SerialNumber = coFreedomDevice.SerialNumber;
            model.Model = coFreedomDevice.Model;
            model.ScheduleNumber = coFreedomDevice.ScheduleNumber;
            model.Exhibit = coFreedomDevice.OwnershipType;
            model.Location = coFreedomDevice.LocName;
            model.LocationID = coFreedomDevice.LocationID;
            model.User = coFreedomDevice.AssetUser;
            model.CostCenter = coFreedomDevice.CostCenter;
            model.Status = coFreedomDevice.Active ;
            model.MonthlyCost = MontlyCost;
            model.InvoiceInstance = instance;
            model.DeviceType = coFreedomDevice.ModelCategory;
            model.TaxRate = taxRate;
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
        public int?   LocationID { get; set; }
        public string User { get; set; }
        public string CostCenter { get; set; }
        public decimal MonthlyCost { get; set; }
        public bool Status { get; set; }
        public string Disposition { get; set; }
        public DateTime? RemovedDateTime { get; set; }
        public string DeviceType {get; set; }
        public decimal TaxRate { get; set; }
        public decimal InvoiceInstance { get; set; }
        public decimal CalculatedTax => MonthlyCost * (TaxRate / 100);
        public decimal UnitTotal => CalculatedTax + MonthlyCost;
    }
}