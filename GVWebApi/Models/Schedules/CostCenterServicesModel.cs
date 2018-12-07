using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
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
}