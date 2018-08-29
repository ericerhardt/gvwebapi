using System;
using System.Linq;
using System.Web.Http;
using GVWebapi.RemoteData;
using GVWebapi.Models;

namespace GVWebapi.Controllers
{
    public class DevicesController : ApiController
    {
        private readonly CoFreedomEntities _coFreedomEntities = new CoFreedomEntities();
        private readonly GlobalViewEntities _globalViewEntities = new GlobalViewEntities(); 

        [HttpGet,Route("api/devices")]
        public IHttpActionResult Get()
        {
            var device = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.ToList().Where(d=> d.Active).Take(100);
            return Ok(device);
        }
        
        [HttpGet,Route("api/devices/{FprID}")]
        public IHttpActionResult GetDevice(string fprId)
        {
            var modelView = new EquipmentsModelView();
            modelView.OpenCalls = 0;
            modelView.equipments = _coFreedomEntities.vw_GVDeviceAnalyzer.FirstOrDefault(c => c.EquipmentNumber == fprId);
            modelView.servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.EquipmentNumber == fprId && !c.Description.Contains("Supply Order")).OrderByDescending(c => c.Date).ToList();
            modelView.orderhistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.EquipmentNumber == fprId && c.Description.Contains("Supply Order")).OrderByDescending(c => c.Date).ToList();
            modelView.OpenCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.EquipmentNumber == fprId && c.v_Status == "Pending");
            var equipment = _globalViewEntities.EquipmentManagers.FirstOrDefault(m => m.Model.Contains(modelView.equipments.Model));

            if(equipment != null)
            {
                modelView.ModelIntro = equipment.IntroDate.GetValueOrDefault();
                modelView.RecoVolume = equipment.MFRMoVol.GetValueOrDefault();
            }
           
            if (modelView.equipments == null)
            {
                return NotFound();
            }
            return Json(modelView);
        }

        [HttpGet,Route("api/customerdevices/{CustomerID}")]
        public IHttpActionResult GetCustomerDevices(int customerId)
        {
            var modelView = new EquipmentsModelView();
            modelView.allequipments = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.ToList().Where(d => d.CustomerID == customerId);
            modelView.activedevices = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Count(d => d.CustomerID == customerId && d.Active);
            modelView.inactivedevices = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Count(d => d.CustomerID == customerId && d.Active == false);
            return Json(modelView);
        }
        
        [HttpGet,Route("api/alldevices/{CustomerID}")]
        public IHttpActionResult GetAllDevices(int customerId)
        {
            var modelView = new EquipmentsModelView();
            modelView.allequipments = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Where(d => d.CustomerID == customerId).ToList();
            modelView.activedevices = modelView.allequipments.Count(e => e.Active);
            modelView.inactivedevices = modelView.allequipments.Count(e => e.Active == false);
            return Json(modelView);
        }

        [HttpGet,Route("api/devicevolumes/{CustomerId}/{Id}/{FromDate?}/{ToDate?}")]
        public IHttpActionResult GetDeviceVolumes(int customerId, string id,DateTime? fromDate, DateTime? toDate)
        {
            fromDate = fromDate ?? DateTime.Now.AddMonths(-11);
            toDate = toDate ?? DateTime.Now;
           
                        var bwdevices = _coFreedomEntities.csDeviceVolumes(fromDate,toDate ,id,customerId)
                            .Where(c => c.MeterType == @"B\W")
                            .Where(c => c.ReadingDate.HasValue)
                            .Select(c=> new { Month = c.ReadingDate.Value.ToString("MMM") , c.Volume}).ToList();

                       var colordevices  = _coFreedomEntities.csDeviceVolumes(fromDate, toDate, id, customerId)
                           .Where(c => c.MeterType == "Color")
                           .Where(c => c.ReadingDate.HasValue)
                           .Select(c => new { Month = c.ReadingDate.Value.ToString("MMM"), c.Volume }).ToList();
             
            var ret = new[]
           {
                   new { label= "B/W Volumes", color = "#768294", data =   bwdevices.Select(c => new object[]{ c.Month ,c.Volume ?? 0m  }) },
                   new { label= "Color Volumes", color = "#1f92fe" , data =  colordevices.Select(c => new object[]{ c.Month ,  c.Volume ?? 0m }) },
               
            };
            return Json(ret);
     
        }

        [HttpGet,Route("api/devicevolumesgrid/{CustomerId}/{Id}/{FromDate?}/{ToDate?}")]
        public IHttpActionResult GetDeviceVolumesGrid(int customerId, string id, DateTime? fromDate, DateTime? toDate)
        {
            var _fromDate = fromDate == null ? DateTime.Now.AddMonths(-11) : fromDate;
            var _toDate = toDate == null ? DateTime.Now : toDate;

            var devices = _coFreedomEntities.csDeviceVolumesGrid(_fromDate, _toDate, id, customerId).ToList();
            var maxbwvolume = _coFreedomEntities.csDeviceVolumesGrid(_fromDate, _toDate, id, customerId).Select(d => d.BWVolume).Max();
            var maxclrvolume = _coFreedomEntities.csDeviceVolumesGrid(_fromDate, _toDate, id, customerId).Select(d => d.ColorVolume).Max();
            
            var maxvolume = maxbwvolume > (maxclrvolume == null ? 0 : maxclrvolume) ? maxbwvolume : maxclrvolume;


            return Json( new { Devices = devices, MaxVolume = maxvolume});

        }
        
        [HttpGet,Route("api/replacementdevices/{CustomerID}")]
        public IHttpActionResult GetReplacementDevices(int customerId)
        {
            var modelView = new EquipmentsModelView();
            modelView.allassetreplacements = _globalViewEntities.AssetReplacements.ToList().Where(d => d.CustomerID == customerId);
            modelView.ReplacementTotals = modelView.allassetreplacements.Sum(s => s.ReplacementValue);
            return Json(modelView);
        }
        
        [HttpGet,Route("api/getreplacement/{replacementid}")]
        public IHttpActionResult GetReplacementDevice(int replacementid)
        {
            var modelView = _globalViewEntities.AssetReplacements.Find(replacementid);
            if (modelView == null) return NotFound();
            return Json(modelView);
           
        }
        
        [HttpGet,Route("api/equipmentmodels/")]
        public IHttpActionResult GetEquipmentModels()
        {
            var models = _coFreedomEntities.vw_EquipmentModels.ToList();
            return Json(models);
        }
        
        [HttpPost,Route("api/addreplacement/")]
        public IHttpActionResult AddReplacement(AssetReplacement replacment)
        {
            if (replacment.CustomerID != null)
            {
                var replacementDevice = new AssetReplacement();
                replacementDevice.CustomerID = replacment.CustomerID;
                replacementDevice.OldModel = replacment.OldModel;
                replacementDevice.OldSerialNumber = replacment.OldSerialNumber;
                replacementDevice.NewModel = replacment.NewModel;
                replacementDevice.NewSerialNumber = replacment.NewSerialNumber;
                replacementDevice.ReplacementDate = replacment.ReplacementDate;
                replacementDevice.Location =   _coFreedomEntities.vw_CustomersOnContract.Where(c => c.CustomerID == replacment.CustomerID).Select(c => c.CustomerName).FirstOrDefault();
                replacementDevice.ReplacementValue = replacment.ReplacementValue;
                _globalViewEntities.AssetReplacements.Add(replacementDevice);
                _globalViewEntities.SaveChanges();

                var modelView = new EquipmentsModelView();
                modelView.allassetreplacements = _globalViewEntities.AssetReplacements.ToList().Where(d => d.CustomerID == replacment.CustomerID);
                return Json(modelView);
                
            }
            return Ok("Device not saved.");
        }
        
        [HttpPost,Route("api/editreplacement/")]
        public IHttpActionResult EditReplacement(AssetReplacement replacment)
        {
            if (replacment.CustomerID != null)
            {
                var replacementDevice = _globalViewEntities.AssetReplacements.Find(replacment.ReplacementID);
                if (replacementDevice == null) return Ok("Device not saved");
                replacementDevice.CustomerID = replacment.CustomerID;
                replacementDevice.OldModel = replacment.OldModel;
                replacementDevice.OldSerialNumber = replacment.OldSerialNumber;
                replacementDevice.NewModel = replacment.NewModel;
                replacementDevice.NewSerialNumber = replacment.NewSerialNumber;
                replacementDevice.ReplacementDate = replacment.ReplacementDate;
                replacementDevice.Location = _coFreedomEntities.vw_CustomersOnContract.Where(c => c.CustomerID == replacment.CustomerID).Select(c => c.CustomerName).FirstOrDefault();
                replacementDevice.ReplacementValue = replacment.ReplacementValue;
                _globalViewEntities.SaveChanges();

                var modelView = new EquipmentsModelView();
                modelView.allassetreplacements = _globalViewEntities.AssetReplacements.ToList().Where(d => d.CustomerID == replacment.CustomerID);
                return Json(modelView);

            }
            return Ok("Device not saved.");
        }
        
        [HttpPost,Route("api/deletereplacement/{id}")]
        public IHttpActionResult DeleteReplacement(int? id)
        {
            if (id != null)
            {
                var replacementDevice = _globalViewEntities.AssetReplacements.Find(id);
                if (replacementDevice == null) return NotFound();
                _globalViewEntities.AssetReplacements.Remove(replacementDevice);
                _globalViewEntities.SaveChanges();
                return Ok("Device Deleted");
            }
            return Ok("Device not Deleted.");
        }
        [HttpGet, Route("api/getdevicesonschedule/{scheduleId}")]
        public IHttpActionResult GetDevicesOnSchedule(long scheduleId)
        {
            var schedule = _globalViewEntities.Schedules.FirstOrDefault(x => x.ScheduleId == scheduleId);
            var devices = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.ScheduleNumber == schedule.Name && x.CustomerID == schedule.CustomerId).ToList();
            return Json(devices);
        }

        [HttpPost,Route("api/updateproperty/{id}/{property}/{value?}")]
        public IHttpActionResult UpdateProperty(int? id, int? property,string value)
        {
            if (id != null && property != null)
            {
                var device = _coFreedomEntities.SCEquipmentCustomProperties.FirstOrDefault(p => p.EquipmentID == id && p.ShAttributeID == property);
                if (device == null) return NotFound();

                if (value == "Not Set")
                {
                    device.TextVal = null;
                }
                else
                {
                    device.TextVal = value;
                }
               
                _coFreedomEntities.SaveChanges();
                return Ok("Device Updated");
            }
            return Ok("Device not Updated.");
        }
        
        [HttpGet,Route("api/servicecalls/{CustomerID}/{StartDate}/{EndDate}/{Type}/{Status}")]
        public IHttpActionResult GetCustomerservicecalls(int customerId,DateTime startDate,DateTime endDate,string type,string status)
        {
            var modelView = new EquipmentsModelView();
            modelView.servicehistories = null;
            modelView.ClosedCalls = 0;
            modelView.OpenCalls = 0;
            modelView.TotalDevices = 0;
//            modelView.EndDate = EndDate.AddDays(1);
            if (type == "All")
            {
              switch (status)
                {
                 case   "All" :
                        modelView.servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled")).OrderByDescending(c => c.Date).ToList();
                         modelView.ClosedCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled"));
                        modelView.OpenCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "Pending"));

                        break;
                    case "Completed":
                         modelView.servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == status)).OrderByDescending(c => c.Date).ToList();
                         modelView.ClosedCalls =   _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled"));
                        modelView.OpenCalls = 0; 
                        break;
                    case "Pending":
                        modelView.servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == status)).OrderByDescending(c => c.Date).ToList();
                        modelView.OpenCalls =   _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "Pending"));
                        modelView.ClosedCalls = 0; 
                        break;
                    case "None":
                        modelView.servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "None")).OrderByDescending(c => c.Date).ToList(); 
                        modelView.ClosedCalls = 0;
                        modelView.OpenCalls = 0;
                        break;
                }



                modelView.TotalDevices = _coFreedomEntities.vw_admin_SCEquipments_22.Count(c => c.CustomerID == customerId & c.Active);

            }
            else
            {
                switch (status)
                {
                    case "All":
                        modelView.servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled" && c.Type == type)).OrderByDescending(c => c.Date).ToList();
                        modelView.ClosedCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled" && c.Type == type));
                        modelView.OpenCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "Pending" && c.Type == type));

                        break;
                    case "Completed":
                        modelView.servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == status && c.Type == type)).OrderByDescending(c => c.Date).ToList();
                        modelView.ClosedCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled" && c.Type == type));
                        modelView.OpenCalls = 0; 
                        break;
                    case "Pending":
                        modelView.servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == status && c.Type == type)).OrderByDescending(c => c.Date).ToList();
                        modelView.OpenCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "Pending" && c.Type == type));
                        modelView.ClosedCalls = 0; 
                        break;
                    case "None":
                        modelView.servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "None")).OrderByDescending(c => c.Date).ToList();
                        modelView.ClosedCalls = 0;
                        modelView.OpenCalls = 0;
                        break;
                }
                      modelView.TotalDevices = _coFreedomEntities.vw_admin_SCEquipments_22.Where(c => c.CustomerID == customerId & c.Active == true).Count();

            }
           
            return Json(modelView);
        }
    }
}
