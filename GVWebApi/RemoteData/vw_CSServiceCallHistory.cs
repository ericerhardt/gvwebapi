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
    
    public partial class vw_CSServiceCallHistory
    {
        public int CallID { get; set; }
        public System.DateTime Date { get; set; }
        public string Caller { get; set; }
        public string EquipmentNumber { get; set; }
        public string Model { get; set; }
        public string Location { get; set; }
        public Nullable<decimal> EquipmentAvgMonthlyVolume { get; set; }
        public string Description { get; set; }
        public string WorkOrderRemarks { get; set; }
        public string CustomerName { get; set; }
        public string CallNumber { get; set; }
        public string WorkOrderNumber { get; set; }
        public int CustomerID { get; set; }
        public string CustomerNumber { get; set; }
        public string Status { get; set; }
        public Nullable<System.DateTime> CloseDate { get; set; }
        public Nullable<System.DateTime> EstStartDate { get; set; }
        public Nullable<int> Duration { get; set; }
        public string v_Status { get; set; }
        public string Type { get; set; }
        public Nullable<int> OpenFor { get; set; }
    }
}
