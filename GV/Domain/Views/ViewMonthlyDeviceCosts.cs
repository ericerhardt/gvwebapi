using System;

namespace GV.Domain.Views
{
    public class ViewMonthlyDeviceCosts
    {
        public int InvoiceId { get; set; }
        public string CustomerName { get; set; }
        public string ContractNumber { get; set; }
        public string EquipmentCustomerName { get; set; }
        public string ContractCode { get; set; }
        public string EquipmentNumber { get; set; }
        public string EquipmentSerialNumber { get; set; }
        public string ContractMeterGroup { get; set; }
        public string MeterType { get; set; }
        public string InvoiceNumber { get; set; }
        public DateTime? BeginMeterDate { get; set; }
        public decimal BeginMeterActual { get; set; }
        public DateTime? EndMeterDate { get; set; }
        public decimal EndMeterActual { get; set; }
        public decimal DifferenceCopies { get; set; }
        public decimal EffectiveRate { get; set; }
        public int CustomerId { get; set; }
        public decimal BilledAmount { get; set; }
    }
}