//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace GVWebapi.RemoteData
{
    using System;
    using System.Collections.Generic;
    
    public partial class vw_admin_SCBillingMeters_MeterGroup
    {
        public int InvoiceID { get; set; }
        public string CustomerName { get; set; }
        public string ContractNumber { get; set; }
        public string EquipmentCustomerName { get; set; }
        public string ContractCode { get; set; }
        public string EquipmentNumber { get; set; }
        public string EquipmentSerialNumber { get; set; }
        public string ContractMeterGroup { get; set; }
        public string MeterType { get; set; }
        public string InvoiceNumber { get; set; }
        public string CreatorID { get; set; }
        public Nullable<System.DateTime> CreateDate { get; set; }
        public Nullable<System.DateTime> BeginMeterDate { get; set; }
        public Nullable<decimal> BeginMeterActual { get; set; }
        public Nullable<System.DateTime> EndMeterDate { get; set; }
        public Nullable<decimal> EndMeterActual { get; set; }
        public Nullable<decimal> DifferenceCopies { get; set; }
        public int EquipmentID { get; set; }
    }
}
