using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using GVWebapi.Models.CostAllocation;
using GVWebapi.RemoteData;
 

namespace GVWebapi.Models.Schedules
{
    public class CostCenterServicesModel
    {
        public string ScheduleName { get; set; }
        public DateTime ScheduleStartDate { get; set; }
        public IEnumerable<CostCenterService> CostCenterServices { get; set; }
    }
    public class CCServicesSummaryModel
    {
        public string MeterGroup { get; set; }
        public Int32 ContractedPages { get; set; }
        public Decimal BaseCPP { get; set; }
        public Decimal OverageCPP { get; set; }
        public Decimal Cost { get; set; }
    }
    
    public class  ServiceCostCenterViewModel
    {
        public int ScheduleCostCenterID { get; set; }
        public long ScheduleID { get; set; }
        public long CustomerID { get; set; }
        public string CostCenter { get; set; }
        public Nullable<int> Status { get; set; }
        public List<MeterGroupCostCenter> MeterGroups { get; set; } = new List<MeterGroupCostCenter>();
    }
    public class MeterGroupCostCenter 
    {
        public string MeterGroupDesc { get; set; }
        public long ContractMeterGroupID { get; set; }
        public int ScheduleServiceID { get; set; }
        public Nullable<int> Volume { get; set; }
        public Nullable<decimal> ExcessCPP { get; set; }
        public decimal InstanceInvoiced { get; set; }
    }

    public  class ScheduleCostCenterViewModel
    {
        public int ScheduleCostCenterID { get; set; }
        public long ScheduleID { get; set; }
        public long CustomerID { get; set; }
        public string CostCenter { get; set; }
        public string MeterGroupDesc { get; set; }
        public Nullable<long> MeterGroupID { get; set; }
        public Nullable<int> Volume { get; set; }
    }
    public class CostCenterModel
    {
        public string MeterGroupDesc { get; set; }
        public long ContractMeterGroupID { get; set; }
        public long ScheduleID { get; set; }
        public string ScheduleName { get; set; }
        public string CostCenter { get; set; }
        public Nullable<int> Volume { get; set; }
        public Nullable<decimal> ExcessCPP { get; set; }
        public Nullable<decimal> BaseCPP { get; set; }
        public Nullable<decimal> TaxRate { get; set; }
        public decimal InstanceInvoiced { get; set; }
        public bool Removed { get; set; }
    }


    public class PeriodAllocatedServicesModel
    {
     public   List<AllocatedServicesViewModel> AllocatedServices { get; set; }
     public   List<MeterGroup> MeterGroups { get; set; }
    }
}