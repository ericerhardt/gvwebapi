using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class VolumeTrendModel
    {
        public long LineID { get; set; }
        public string DeviceID { get; set; }
        public string Model { get; set; }
        public string MeterGroup { get; set; }
        public string SerialNumber { get; set; }
        public string DeviceStatus { get; set; }
        public string Location { get; set; }
        public string Building { get; set; }
        public string CostCenter { get; set; }
        public string Dept { get; set; }
        public string Floor { get; set; }
        public string User { get; set; }
        public string Comments { get; set; }
        public string IPAddress { get; set; }
        public Nullable<System.DateTime> StartDate { get; set; }
        public Nullable<System.DateTime> EndDate { get; set; }
        public Nullable<decimal> StartMeter { get; set; }
        public Nullable<decimal> EndMeter { get; set; }
        public string DeviceType { get; set; }
        public Nullable<decimal> PercOfTotal { get; set; }
        public Nullable<decimal> PercOfMeterGroup { get; set; }
        public Nullable<decimal> LastPeriodVolume { get; set; }
        public Nullable<decimal> PeriodVolume { get; set; }
        public Nullable<decimal> VolumeDiff { get; set; }
        public Nullable<decimal> TotalCopies { get; set; }
        public Nullable<decimal> StartMeterActual { get; set; }
        public Nullable<decimal> EndMeterActual { get; set; }
    }
}