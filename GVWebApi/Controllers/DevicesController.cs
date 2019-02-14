using System;
using System.Linq;
using System.Web.Http;
using GVWebapi.RemoteData;
using GVWebapi.Models;
using GVWebapi.Helpers;
using System.Data.Entity.Core.Objects;
using System.Data.Common;
using System.Data;
using GVWebapi.Services;
using System.Data.Entity.SqlServer;
using System.Web.UI.WebControls;
using System.Collections.Generic;

namespace GVWebapi.Controllers
{
    public class DevicesController : ApiController  
    {
        private readonly CoFreedomEntities _coFreedomEntities = new CoFreedomEntities();
        private readonly GlobalViewEntities _globalViewEntities = new GlobalViewEntities();

        [HttpGet, Route("api/devices")]
        public IHttpActionResult Get()
        {
            var device = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.ToList().Where(d => d.Active).Take(100);
            return Ok(device);
        }

        [HttpGet, Route("api/devices/{FprID}")]
        public IHttpActionResult GetDevice(string fprId)
        {
            var modelView = new EquipmentsModelView();
            modelView.OpenCalls = 0;
            modelView.Equipments = _coFreedomEntities.vw_GVDeviceAnalyzer.FirstOrDefault(c => c.EquipmentNumber == fprId);
            modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.EquipmentNumber == fprId && !c.Description.Contains("Supply Order")).OrderByDescending(c => c.Date).ToList();
            modelView.Orderhistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.EquipmentNumber == fprId && c.Description.Contains("Supply Order")).OrderByDescending(c => c.Date).ToList();
            modelView.OpenCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.EquipmentNumber == fprId && c.v_Status == "Pending");
            modelView.PeriodStartList = _coFreedomEntities.vw_REVisionInvoices.Where(x => x.ContractID == modelView.Equipments.ContractID && x.StartDate >= modelView.Equipments.InstallDate ).Select( x => new DateListModel {DateValue = x.StartDate.Value, DateString =  x.StartDateString }).ToList();
            
            modelView.PeriodEndList = _coFreedomEntities.vw_REVisionInvoices.Where(x => x.ContractID == modelView.Equipments.ContractID && x.PeriodDate >= modelView.Equipments.InstallDate).OrderByDescending(x=> x.PeriodDate).Select(x => new DateListModel { DateValue = x.PeriodDate.Value, DateString = x.PeriodDateString }).ToList();
            if (modelView.PeriodStartList.Count() == 0)
            {
                modelView.PeriodStartList = modelView.PeriodEndList;
            }
            var equipment = _globalViewEntities.EquipmentManagers.FirstOrDefault(m => m.Model.Contains(modelView.Equipments.Model));

            if (equipment != null)
            {
                modelView.ModelIntro = equipment.IntroDate.GetValueOrDefault();
                modelView.RecoVolume = equipment.MFRMoVol.GetValueOrDefault();
            }

            if (modelView.Equipments == null)
            {
                return NotFound();
            }
            return Json(  modelView );
        }
        [HttpGet, Route("api/mobiledevices/{FprID}")]
        public IHttpActionResult GetMobileDevice(string fprId)
        {
            var modelView = new EquipmentsModelView();
            modelView.OpenCalls = 0;
            modelView.Equipments = _coFreedomEntities.vw_GVDeviceAnalyzer.FirstOrDefault(c => c.EquipmentNumber == fprId);
            modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.EquipmentNumber == fprId).OrderByDescending(c => c.Date).Take(3).ToList();
            var equipment = _globalViewEntities.EquipmentManagers.FirstOrDefault(m => m.Model.Contains(modelView.Equipments.Model));

            if (equipment != null)
            {
                modelView.ModelIntro = equipment.IntroDate.GetValueOrDefault();
                modelView.RecoVolume = equipment.MFRMoVol.GetValueOrDefault();
            }

            if (modelView.Equipments == null)
            {
                return NotFound();
            }
            return Json(new { equipments = modelView.Equipments, servicehistories = modelView.Servicehistories });
        }

        [HttpGet, Route("api/customerdevices/{CustomerID}")]
        public IHttpActionResult GetCustomerDevices(int customerId)
        {
            var modelView = new EquipmentsModelView();
            modelView.AllEquipments = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.ToList().Where(d => d.CustomerID == customerId);
            modelView.activedevices = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Count(d => d.CustomerID == customerId && d.Active);
            modelView.inactivedevices = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Count(d => d.CustomerID == customerId && d.Active == false);
            return Json(modelView);
        }

        [HttpGet, Route("api/alldevices/{CustomerID}")]
        public IHttpActionResult GetAllDevices(int customerId)
        {
            var modelView = new EquipmentsModelView();
            modelView.AllEquipments = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Where(d => d.CustomerID == customerId).ToList();
            modelView.activedevices = modelView.AllEquipments.Count(e => e.Active);
            modelView.inactivedevices = modelView.AllEquipments.Count(e => e.Active == false);
            return Json(modelView);
        }

        [HttpGet, Route("api/devicevolumes/{CustomerId}/{Id}/{FromDate?}/{ToDate?}")]
        public IHttpActionResult GetDeviceVolumes(int customerId, string id, DateTime? fromDate, DateTime? toDate)
        {
            fromDate = fromDate ?? DateTime.Now.AddMonths(-11);
            toDate = toDate ?? DateTime.Now;

            var bwdevices = _coFreedomEntities.csDeviceVolumes(fromDate, toDate, id, customerId)
                .Where(c => c.MeterType == @"B\W")
                .Where(c => c.ReadingDate.HasValue)
                .Select(c => new { Month = c.ReadingDate.Value.ToString("MMM YY"), c.Volume }).ToList();

            var colordevices = _coFreedomEntities.csDeviceVolumes(fromDate, toDate, id, customerId)
                .Where(c => c.MeterType == "Color")
                .Where(c => c.ReadingDate.HasValue)
                .Select(c => new { Month = c.ReadingDate.Value.ToString("MMM YY"), c.Volume }).ToList();

            var ret = new[]
           {
                   new { label= "B/W Volumes", color = "#768294", data =   bwdevices.Select(c => new object[]{ c.Month ,c.Volume ?? 0m  }) },
                   new { label= "Color Volumes", color = "#1f92fe" , data =  colordevices.Select(c => new object[]{ c.Month ,  c.Volume ?? 0m }) },

            };
            return Json(ret);

        }

        [HttpGet, Route("api/devicevolumesgrid/{CustomerId}/{Id}/{FromDate?}/{ToDate?}")]
        public IHttpActionResult GetDeviceVolumesGrid(int customerId, string id, DateTime? fromDate, DateTime? toDate)
        {
            var _fromDate = fromDate == null ? DateTime.Now.AddMonths(-11) : fromDate;
            var _toDate = toDate == null ? DateTime.Now : toDate;

            var devices = _coFreedomEntities.csDeviceVolumesGrid(_fromDate, _toDate, id, customerId).ToList();
            var maxbwvolume = _coFreedomEntities.csDeviceVolumesGrid(_fromDate, _toDate, id, customerId).Select(d => d.BWVolume).Max();
            var maxclrvolume = _coFreedomEntities.csDeviceVolumesGrid(_fromDate, _toDate, id, customerId).Select(d => d.ColorVolume).Max();

            var maxvolume = maxbwvolume > (maxclrvolume == null ? 0 : maxclrvolume) ? maxbwvolume : maxclrvolume;


            return Json(new { Devices = devices, MaxVolume = maxvolume });

        }
        [HttpGet, Route("api/devicevolumesgridv2/{CustomerId}/{Id}/{FromDate?}/{ToDate?}")]
        public IHttpActionResult GetDeviceVolumesGridV2(int customerId, string id, DateTime? fromDate, DateTime? toDate)
        {
            var _fromDate = fromDate == null ? DateTime.Now.AddMonths(-11) : fromDate;
            var _toDate = toDate == null ? DateTime.Now : toDate;


            var devices = _coFreedomEntities.vw_BillingMeterVolumesByDevice
                                                    .Where(c => c.EquipmentNumber == id)
                                                    .Where(c => c.ReadingDate >= fromDate && c.ReadingDate <= toDate).ToList();
            
            decimal? maxbwvolume = devices.Max(x=> x.BWVolume);
            decimal? maxclrvolume = devices.Max(x=> x.ColorVolume);

            var maxvolume = maxbwvolume > (maxclrvolume == null ? 0 : maxclrvolume) ? maxbwvolume : maxclrvolume;


            return Json(new { Devices = devices, MaxVolume = maxvolume });

        }



        [HttpGet, Route("api/devicevolumesv2/{CustomerId}/{Id}/{FromDate?}/{ToDate?}")]
        public IHttpActionResult GetDeviceVolumesV2(int customerId, string id, DateTime? fromDate, DateTime? toDate)
        {
            fromDate = fromDate ?? DateTime.Now.AddMonths(-11);
            toDate = toDate ?? DateTime.Now;

            List<vw_BillingMeterVolumesByDevice> devices = _coFreedomEntities.vw_BillingMeterVolumesByDevice
                                                  .Where(c => c.EquipmentNumber == id)
                                                  .Where(c => c.ReadingDate >= fromDate && c.ReadingDate <= toDate).ToList();

            

            var ret = new[]
           {
                   new { label= "B/W Volumes", color = "#768294", data =   devices.Select(c => new object[]{  DateTimeExtensions.GetJavascriptTimeStamp(c.ReadingDate.Value)   ,c.BWVolume.Value.ToString("#.00")  }) },
                   new { label= "Color Volumes", color = "#1f92fe" , data =  devices.Select(c => new object[]{ DateTimeExtensions.GetJavascriptTimeStamp(c.ReadingDate.Value) ,  c.ColorVolume.ToString("#.00")   }) },

            };
            return Json(ret);

        }
        
        [HttpGet, Route("api/replacementdevices/{CustomerID}")]
        public IHttpActionResult GetReplacementDevices(int customerId)
        {
            var modelView = new EquipmentsModelView();
            modelView.Allassetreplacements = _globalViewEntities.AssetReplacements.ToList().Where(d => d.CustomerID == customerId);
            modelView.ReplacementTotals = modelView.Allassetreplacements.Sum(s => s.ReplacementValue);
            return Json(modelView);
        }

        [HttpGet, Route("api/getreplacement/{replacementid}")]
        public IHttpActionResult GetReplacementDevice(int replacementid)
        {
            var modelView = _globalViewEntities.AssetReplacements.Find(replacementid);
            if (modelView == null) return NotFound();
            return Json(modelView);

        }

        [HttpGet, Route("api/equipmentmodels/")]
        public IHttpActionResult GetEquipmentModels()
        {
            var models = _coFreedomEntities.vw_EquipmentModels.ToList();
            return Json(models);
        }

        [HttpPost, Route("api/addreplacement/")]
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
                replacementDevice.Comments = replacment.Comments;
                replacementDevice.Location = _coFreedomEntities.vw_CustomersOnContract.Where(c => c.CustomerID == replacment.CustomerID).Select(c => c.CustomerName).FirstOrDefault();
                replacementDevice.ReplacementValue = replacment.ReplacementValue;
                _globalViewEntities.AssetReplacements.Add(replacementDevice);
                _globalViewEntities.SaveChanges();

                var modelView = new EquipmentsModelView();
                modelView.Allassetreplacements = _globalViewEntities.AssetReplacements.ToList().Where(d => d.CustomerID == replacment.CustomerID);
                return Json(modelView);

            }
            return Ok("Device not saved.");
        }

        [HttpPost, Route("api/editreplacement/")]
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
                modelView.Allassetreplacements = _globalViewEntities.AssetReplacements.ToList().Where(d => d.CustomerID == replacment.CustomerID);
                return Json(modelView);

            }
            return Ok("Device not saved.");
        }

        [HttpPost, Route("api/deletereplacement/{id}")]
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
            var schedule = _globalViewEntities.Schedules.Find(scheduleId);
            var devices = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.ScheduleNumber == schedule.Name && x.CustomerID == schedule.CustomerId && x.NumberOfContractsActive > 0).ToList();
            schedule.MonthlyHwCost = devices.Select(i => Decimal.Parse(i.MonthlyCost)).Sum();
            _globalViewEntities.SaveChanges();
            return Json(devices);
        }
        [HttpGet, Route("api/getdevicesremoved/{customerId}")]
        public IHttpActionResult GetDevicesRemoved(int customerId)
        {

            var devices = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => x.CustomerID == customerId && x.NumberOfContractsActive == null).ToList();
            return Json(devices);
        }
        [HttpGet, Route("api/getdevicesunallocated/{customerId}")]
        public IHttpActionResult GetDevicesUnallocated(int customerId)
        {
            var schedules = _globalViewEntities.Schedules.Where(x => x.CustomerId == customerId && x.IsDeleted == false).Select(x => x.Name).ToList();
            var scheduleList = _globalViewEntities.Schedules.Where(x => x.CustomerId == customerId && x.IsDeleted == false).Select(x => new { Name = x.Name, ScheduleID = x.ScheduleId }).ToList();
            var devices = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.Where(x => (x.CustomerID == customerId && x.NumberOfContractsActive > 0) && !schedules.Contains(x.ScheduleNumber)).ToList();
            return Json(new { devices, schedules = scheduleList });
        }
        [HttpPost, Route("api/updateproperty/{id}/{property}/{value?}")]
        public IHttpActionResult UpdateProperty(int? id, int? property, string value)
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
        [HttpPost, Route("api/placeservicecall/")]
        public IHttpActionResult PlaceServiceCall(ServiceCallModel model)
        {
            
            if (model != null)
            {

                if (model.isWorking)
                {
                    model.Description = model.Description + " This device is functioning";
                }
                else
                {
                    model.Description = model.Description + " This device is not functioning";
                }

                var Id = InsertServiceCall(model);
                
               var success =  MailParser.EmailServiceCall(Id, model, 1);
                if (success)
                {
                    return Json(new { status = "submitted", results = model });
                }  
                
            }

            return Json(new { status = "error", results = BadRequest() });
        }
        [HttpPost, Route("api/webservicecall/")]
        public IHttpActionResult WebServiceCall(ServiceCallModel model)
        {
            var CallId = "0";
            if (model != null)
            {
                if (model.EquipmentID == 0)
                {
                    var equip = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.FirstOrDefault(x => x.EquipmentNumber == model.EquipmentNumber);
                    if (equip != null)
                    {
                        model.EquipmentID = equip.EquipmentID;
                        if (model.isWorking)
                        {
                            model.Description = model.Description + " This device is functioning";
                        }
                        else
                        {
                            model.Description = model.Description + " This device is not functioning";
                        }

                        CallId = InsertServiceCall(model);
                        var success = MailParser.EmailServiceCall(CallId, model, 1);
                        if (success)
                        {

                            return Redirect("https://www.fprus.com/servicesupply_ty");
                        }
                        
                    } else
                    {
                        var success = MailParser.EmailServiceCall(CallId, model, 3);
                        if (success)
                        {

                            return Redirect("https://www.fprus.com/servicesupply_ty");
                        }
                       
                    }
                     
                }

            }

            return Redirect("https://www.fprus.com/service_call_error");
        }
        [HttpPost, Route("api/integrisservicecall/")]
        public IHttpActionResult IntegrisServiceCall(IntegrisServiceCallModel imodel)
        {
            var CallId = "0";
            if (imodel != null)
            {
                var model = new ServiceCallModel
                {
                    EquipmentNumber = imodel.EquipmentNumber,
                    Name = imodel.Name,
                    Phone = imodel.Phone,
                    Email = imodel.Email,
                    Description = imodel.Description,
                    isWorking = imodel.IsWorking
                };


                if (model.EquipmentID == 0)
                {
                    var equip = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.FirstOrDefault(x => x.EquipmentNumber == model.EquipmentNumber);
                    if (equip != null)
                    {
                        imodel.EquipmentID = equip.EquipmentID;
                        model.EquipmentID = equip.EquipmentID;
                        if (model.isWorking)
                        {
                            model.Description = model.Description + ". This device is functioning.";
                        }
                        else
                        {
                            model.Description = model.Description + ". This device is not functioning.";
                        }
                        if (imodel.CanPrint)
                        {
                            model.Description = model.Description + "  We can not print to another device.";
                        }
                        if (imodel.PatientCare)
                        {
                            model.Description = model.Description + " This is impacting patient care.";
                        }
                        imodel.Description = model.Description;
                        CallId = InsertServiceCall(model);
                        var success = MailParser.EmailIntegrisServiceCall(CallId, imodel, 1);
                        if (success)
                        {
                            return Redirect("https://www.fprus.com/client-integris/integris-servicesupply_ty");
                        }
                    }
                    else
                    {
                        var success = MailParser.EmailIntegrisServiceCall(CallId, imodel, 3);
                        if (success)
                        {
                            return Redirect("https://www.fprus.com/client-integris/integris-servicesupply_ty");
                        }

                    }


                }
 
            }

            return Redirect("https://www.fprus.com/service_call_error");
        }

        [HttpPost, Route("api/placesupplycall/")]
        public IHttpActionResult PlaceSupplyCall(ServiceCallModel model)
        {
             
            if (model != null)
            {
               
                var Id = InsertSupplyCall(model);
                var success = MailParser.EmailServiceCall(Id, model, 2);
                if (success)
                {
                    return Json(new { status = "submitted", results = model });
                }
            }
            return Json(new { status = "error", results = BadRequest() });

        }
        [HttpPost, Route("api/websupplycall/")]
        public IHttpActionResult WebSupplyCall(ServiceCallModel model)
        {
            var CallId = "0";
            if (model != null)
            {
                var equip = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.FirstOrDefault(x => x.EquipmentNumber == model.EquipmentNumber);
                if (equip != null)
                {
                    model.EquipmentID = equip.EquipmentID;
                    CallId = InsertSupplyCall(model);
                    var success = MailParser.EmailServiceCall(CallId, model, 2);
                    if (success)
                    {
                        return Redirect("https://www.fprus.com/servicesupply_ty");
                    }
                    
                } else
                {
                    var success = MailParser.EmailServiceCall(CallId, model, 4);
                    if (success)
                    {
                        return Redirect("https://www.fprus.com/servicesupply_ty");
                    }
                   
                }
                
                  
                
            }
            return Redirect("https://www.fprus.com/service_call_error");

        }
        [HttpPost, Route("api/integrissupplycall/")]
        public IHttpActionResult IntegrisSupplyCall(IntegrisServiceCallModel imodel)
        {
            var CallId = "0";
            if (imodel != null)
            {
                var model = new ServiceCallModel
                {
                    EquipmentNumber = imodel.EquipmentNumber,
                    Name = imodel.Name,
                    Phone = imodel.Phone,
                    Email = imodel.Email,
                    Description = imodel.Description,
                    isWorking = imodel.IsWorking,
                    Black = imodel.Black,
                    Cyan = imodel.Cyan,
                    Magenta = imodel.Magenta,
                    Yellow = imodel.Yellow,


                };

                var equip = _coFreedomEntities.vw_admin_EquipmentList_MeterGroup.FirstOrDefault(x => x.EquipmentNumber == model.EquipmentNumber);
                if (equip != null)
                {
                    imodel.EquipmentID = equip.EquipmentID;
                    model.EquipmentID = equip.EquipmentID;
                    CallId = InsertSupplyCall(model);
                    var success = MailParser.EmailIntegrisServiceCall(CallId, imodel, 2);
                    if (success)
                    {
                        return Redirect(" https://www.fprus.com/client-integris/integris-servicesupply_ty");
                    }

                }
                else
                {
                    var success = MailParser.EmailIntegrisServiceCall(CallId, imodel, 4);
                    if (success)
                    {
                        return Redirect(" https://www.fprus.com/client-integris/integris-servicesupply_ty");
                    }

                }



            }
            return Redirect("https://www.fprus.com/service_call_error");

        }
        [HttpGet, Route("api/servicecalls/{CustomerID}/{StartDate}/{EndDate}/{Type}/{Status}")]
        public IHttpActionResult GetCustomerservicecalls(int customerId, DateTime startDate, DateTime endDate, string type, string status)
        {
            var modelView = new EquipmentsModelView();
            modelView.Servicehistories = null;
            modelView.ClosedCalls = 0;
            modelView.OpenCalls = 0;
            modelView.TotalDevices = 0;
            //            modelView.EndDate = EndDate.AddDays(1);
            if (type == "All")
            {
                switch (status)
                {
                    case "All":
                        modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled")).OrderByDescending(c => c.Date).ToList();
                        modelView.ClosedCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled"));
                        modelView.OpenCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "Pending"));

                        break;
                    case "Completed":
                        modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == status)).OrderByDescending(c => c.Date).ToList();
                        modelView.ClosedCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled"));
                        modelView.OpenCalls = 0;
                        break;
                    case "Pending":
                        modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == status)).OrderByDescending(c => c.Date).ToList();
                        modelView.OpenCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "Pending"));
                        modelView.ClosedCalls = 0;
                        break;
                    case "None":
                        modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "None")).OrderByDescending(c => c.Date).ToList();
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
                        modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled" && c.Type == type)).OrderByDescending(c => c.Date).ToList();
                        modelView.ClosedCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled" && c.Type == type));
                        modelView.OpenCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "Pending" && c.Type == type));

                        break;
                    case "Completed":
                        modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == status && c.Type == type)).OrderByDescending(c => c.Date).ToList();
                        modelView.ClosedCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status != "Canceled" && c.Type == type));
                        modelView.OpenCalls = 0;
                        break;
                    case "Pending":
                        modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == status && c.Type == type)).OrderByDescending(c => c.Date).ToList();
                        modelView.OpenCalls = _coFreedomEntities.vw_CSServiceCallHistory.Count(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "Pending" && c.Type == type));
                        modelView.ClosedCalls = 0;
                        break;
                    case "None":
                        modelView.Servicehistories = _coFreedomEntities.vw_CSServiceCallHistory.Where(c => c.CustomerID == customerId && (c.Date <= endDate && c.Date >= startDate && c.v_Status == "None")).OrderByDescending(c => c.Date).ToList();
                        modelView.ClosedCalls = 0;
                        modelView.OpenCalls = 0;
                        break;
                }
                modelView.TotalDevices = _coFreedomEntities.vw_admin_SCEquipments_22.Where(c => c.CustomerID == customerId & c.Active == true).Count();

            }

            return Json(modelView);
        }
        public static string InsertSupplyCall(ServiceCallModel oSupplyInfo)
        {
            var db = new CoFreedomEntities();
            DbConnection con = db.Database.Connection;
            DbCommand cmd = con.CreateCommand();
            try
            {
                con.Open();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "Web_SCInsertServiceCall";

                DbParameter paramCaller = cmd.CreateParameter();
                paramCaller.ParameterName = "Caller";
                paramCaller.Value = oSupplyInfo.Name +"("+ oSupplyInfo.Phone +"/"+oSupplyInfo.Email+")";
                cmd.Parameters.Add(paramCaller);

                DbParameter paramEquipid = cmd.CreateParameter();
                paramEquipid.ParameterName = "EquipmentID";
                paramEquipid.Value = oSupplyInfo.EquipmentID;
                cmd.Parameters.Add(paramEquipid);

                DbParameter paramDescription = cmd.CreateParameter();
                paramDescription.ParameterName = "Description";
                paramDescription.Value = "Supply Order For Device " + oSupplyInfo.EquipmentNumber + "\r\n Black =" + oSupplyInfo.Black + "\r\n Cyan=" + oSupplyInfo.Cyan + "\r\n Magenta=" + oSupplyInfo.Magenta + "\r\n Yellow=" + oSupplyInfo.Yellow + "\r\n Do they have supplies?:" + oSupplyInfo.isWorking.ToString() + "\r\n Comments:\r\n " + oSupplyInfo.Description;
                cmd.Parameters.Add(paramDescription);

                DbParameter paramCallType = cmd.CreateParameter();
                paramCallType.ParameterName = "CallTypeID";
                paramCallType.Value = 4;
                cmd.Parameters.Add(paramCallType);

                DbParameter paramCallId = cmd.CreateParameter();
                paramCallId.ParameterName = "CallID";
                paramCallId.DbType = DbType.Int32;
                paramCallId.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramCallId);

                cmd.ExecuteNonQuery();

                return cmd.Parameters["CallID"].Value.ToString();
            }
            catch (Exception)
            {
                return "!UNKNOWN!";
            }
            finally
            {
                cmd.Dispose();
                con.Dispose();
            }
        }
        public static string InsertServiceCall(ServiceCallModel oSupplyInfo)
        {
            var db = new CoFreedomEntities();
            DbConnection con = db.Database.Connection;
            DbCommand cmd = con.CreateCommand();
            try
            {
                con.Open();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "Web_SCInsertServiceCall";

                DbParameter paramCaller = cmd.CreateParameter();
                paramCaller.ParameterName = "Caller";
               
                paramCaller.Value = oSupplyInfo.Name + "(" + oSupplyInfo.Phone + "/" + oSupplyInfo.Email + ")";
                cmd.Parameters.Add(paramCaller);

                DbParameter paramEquipid = cmd.CreateParameter();
                paramEquipid.ParameterName = "EquipmentID";
                paramEquipid.Value = oSupplyInfo.EquipmentID;
                cmd.Parameters.Add(paramEquipid);

                DbParameter paramDescription = cmd.CreateParameter();
                paramDescription.ParameterName = "Description";
                paramDescription.Value = oSupplyInfo.Description;
                cmd.Parameters.Add(paramDescription);

                DbParameter paramCallType = cmd.CreateParameter();
                paramCallType.ParameterName = "CallTypeID";
                paramCallType.Value = 4;
                cmd.Parameters.Add(paramCallType);

                DbParameter paramCallId = cmd.CreateParameter();
                paramCallId.ParameterName = "CallID";
                paramCallId.DbType = DbType.Int32;
                paramCallId.Direction = ParameterDirection.Output;
                cmd.Parameters.Add(paramCallId);

                cmd.ExecuteNonQuery();

                return cmd.Parameters["CallID"].Value.ToString();
            }
            catch (Exception ex)
            {
                return "!UNKNOWN!";
            }
            finally
            {
                cmd.Dispose();
                con.Dispose();
            }

        }
    }
 
}
