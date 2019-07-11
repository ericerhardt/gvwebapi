using System;
using System.ComponentModel.DataAnnotations;


namespace GVWebapi.Models.CostAllocation
{
    public class vw_GV_InvoicedEquipmentHistory
    {
        [Key]
        public int InvoiceID { get; set; }
        public int Period { get; set; }
        public DateTime PeriodDate { get; set; }
        public int EquipmentID { get; set; }
        public string EquipmentNumber { get; set; }
        public string Model { get; set; }
        public string Location { get; set; }
        public string SerialNumber { get; set; }
        public string Building { get; set; }
        public string User { get; set; }
        public string Floor { get; set; }
        public string Department { get; set; }
        public string CostCenter { get; set; }
        public string Address { get; set; }
        public string Exhibit { get; set; }
        public string ScheduleNumber { get; set; }
        public string IPAddress { get; set; }
        public int CustomerID { get; set; }
        public int BillToID { get; set; }
        public int LocationID { get; set; }
        public Nullable<bool> Active { get; set; }
        public Nullable<System.DateTime> InstallDate { get; set; }
        public string MonthlyCost { get; set; }
        public string ContractMeterGroup { get; set; }
        public Nullable<int> ContractMeterGroupID { get; set; }
        public DateTime BeginMeterDate { get; set; }
        public decimal BeginMeter { get; set; }
        public DateTime EndMeterDate { get; set; }
        public decimal EndMeter { get; set; }
        public decimal Volume { get; set; }
    }
    public class ReconcileCostCenter
    {
        public static ReconcileCostCenter For(vw_GV_InvoicedEquipmentHistory invoicedEquipment, decimal HardwareTax)
        {
            Decimal.TryParse(invoicedEquipment.MonthlyCost, out decimal MontlyCost);

            var model = new ReconcileCostCenter();
            model.ContractMeterGroup = invoicedEquipment.ContractMeterGroup;
            model.ContractMeterGroupID = invoicedEquipment.ContractMeterGroupID.Value;
            model.CostCenter = invoicedEquipment.CostCenter;
            model.MonthlyCost = MontlyCost;
            model.LocName = invoicedEquipment.Location;
            model.Volume = (int)invoicedEquipment.Volume;
            model.HardwareTax = HardwareTax;
            return model;
        }

        public string CostCenter { get; set; }
        public string LocName { get; set; }
        public decimal MonthlyCost { get; set; }
        public int Volume { get; set; }
        public int ContractMeterGroupID { get; set; }
        public string ContractMeterGroup { get; set; }
        public decimal HardwareTax { get; set; }
        public decimal CalculatedTax => MonthlyCost * (HardwareTax / 100);
    }
}