using System;
using GV.Domain.Entities;
using GVWebapi.RemoteData;
using GVWebapi.Services;

namespace GVWebapi.Models.Devices
{
    public class ScheduleDevicesModel
    {
        public static ScheduleDevicesModel For(decimal taxRate,decimal instance, ScheduleDevicesEntity devicesModel)
        {
           
            var model = new ScheduleDevicesModel();
            model.EquipmentID = devicesModel.EquipmentID;
            model.CustomerID = devicesModel.CustomerID;
            model.EquipmentNumber = devicesModel.EquipmentNumber;
            model.SerialNumber = devicesModel.SerialNumber;
            model.Model = devicesModel.Model;
            model.ScheduleNumber = devicesModel.ScheduleNumber;
            model.Exhibit = devicesModel.OwnershipType;
            model.Location = devicesModel.Location;
            model.LocationID = devicesModel.LocationID;
            model.User = devicesModel.AssetUser;
            model.CostCenter = devicesModel.CostCenter;
            model.Status = devicesModel.Active.Value ;
            model.MonthlyCost = devicesModel.MonthlyCost;
            model.InvoiceInstance = instance;
            model.TaxRate = taxRate;
            return model;
        }

        private static ScheduleDevicesStatusEnum GetDeviceStatus(vw_admin_EquipmentList_MeterGroup coFreedomDevice, ScheduleDevicesEntity deviceEntity)
        {
            if (deviceEntity.RemovedStatus.HasValue && deviceEntity.RemovedStatus.Value == RemovedStatusEnum.SetForRemoval)
            {
                return ScheduleDevicesStatusEnum.Removed;
            }

            switch (coFreedomDevice.Active)
            {
                case true:
                    return ScheduleDevicesStatusEnum.Active;
                case false:
                    return ScheduleDevicesStatusEnum.InActive;
                default:
                    return ScheduleDevicesStatusEnum.Removed;
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

        private ScheduleDevicesModel()
        {
        }

        public long ScheduleDeviceID { get; set; }
        public int EquipmentID { get; set; }
        public long CustomerID { get; set; }
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