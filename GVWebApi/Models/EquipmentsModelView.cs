using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using GVWebapi.RemoteData;
namespace GVWebapi.Models
{
    public class EquipmentsModelView
    {
        public vw_GVDeviceAnalyzer equipments { get; set; }
        public IEnumerable<vw_CSServiceCallHistory> servicehistories { get; set; }
        public IEnumerable<vw_CSServiceCallHistory> orderhistories { get; set; }
        public IEnumerable<vw_admin_EquipmentList_MeterGroup> allequipments { get; set; } 
        public IEnumerable<AssetReplacement> allassetreplacements { get; set; }
        public DateTime? ModelIntro { get; set; }
        public int? RecoVolume { get; set; }
        public int activedevices { get; set; }
        public int inactivedevices { get; set; }
        public int ClosedCalls { get; set; }
        public int OpenCalls { get; set; }
        public int TotalDevices { get; set; }
        public decimal? ReplacementTotals { get; set; }
    }
}