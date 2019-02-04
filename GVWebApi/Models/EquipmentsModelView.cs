using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;
using GVWebapi.RemoteData;
namespace GVWebapi.Models
{
    public class EquipmentsModelView
    {
        public vw_GVDeviceAnalyzer Equipments { get; set; }
        public IEnumerable<vw_CSServiceCallHistory> Servicehistories { get; set; }
        public IEnumerable<vw_CSServiceCallHistory> Orderhistories { get; set; }
        public IEnumerable<vw_admin_EquipmentList_MeterGroup> AllEquipments { get; set; }
        public List<DateListModel> PeriodStartList { get; set; }
        public List<DateListModel> PeriodEndList { get; set; }
        public IEnumerable<AssetReplacement> Allassetreplacements { get; set; }
        public DateTime? ModelIntro { get; set; }
        public int? RecoVolume { get; set; }
        public int activedevices { get; set; }
        public int inactivedevices { get; set; }
        public int ClosedCalls { get; set; }
        public int OpenCalls { get; set; }
        public int TotalDevices { get; set; }
        public decimal? ReplacementTotals { get; set; }
    }
    public class DateListModel
    {
        public DateTime DateValue { get; set; }
        public string   DateString { get; set; }
    }
}