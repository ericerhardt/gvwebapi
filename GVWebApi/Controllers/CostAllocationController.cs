using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Http;
using GVWebapi.RemoteData;
using GVWebapi.Models;
using GVWebapi.Helpers;
using GVWebapi.Models.Schedules;
using GVWebapi.Services;
using GV.Domain;

namespace GVWebapi.Controllers
{
    public class CostAllocationController : ApiController
    {
 
        private readonly ICostAllocationService _costAllocationService;
        private readonly IUnitOfWork _unitOfWork;
        private readonly GlobalViewEntities _globalview = new GlobalViewEntities();


        public CostAllocationController(ICostAllocationService costAllocationService, IUnitOfWork unitOfWork )
        {
            _costAllocationService = costAllocationService;
            _unitOfWork = unitOfWork;
           
        }

        // GET: api/CostAllocation/5
        [HttpGet,Route("api/allocationsettings/{id}")]
        public async Task<IHttpActionResult> GetCostAllocationSetting(int id)
        {
            
           var  costAllocationSetting = await _globalview.CostAllocationSettings.Where(x =>x.CustomerID == id).FirstOrDefaultAsync();
         
            if (costAllocationSetting == null)
            {
                CostAllocationSetting row = new CostAllocationSetting();
                row.CustomerID = id;
                row.AssumedAllocation = true;
                _globalview.CostAllocationSettings.Add(row);
                _globalview.SaveChanges();  
            }
            costAllocationSetting = await _globalview.CostAllocationSettings.Where(x => x.CustomerID == id).FirstOrDefaultAsync();
            var metergroups = _costAllocationService.GetCostAllocationMeterGroups(id).OrderBy(x => x.ContractMeterGroup);

            return Ok( new { settings = costAllocationSetting, metergroups  });
        }
        [HttpPost, Route("api/saveexcessmetergroups")]
        public async Task<IHttpActionResult> SaveMeterGroupExcessCPP(IEnumerable<CostAllocationSettingsViewModel> models )
        {
            foreach(var model in models)
            {
                var metergroup = _globalview.CostAllocationMeterGroups.Find(model.SettingsMeterGroupID);
                {
                    metergroup.ExcessCPP = model.ExcessCPP;
                };
               
            }
            await _globalview.SaveChangesAsync();
        

            return Ok();
        }

        [HttpPost, Route("api/costallocationsave")]
        public async Task<IHttpActionResult> SaveCostAllocationSave(CostAllocationSetting model)
        {
            
            var setting = _globalview.CostAllocationSettings.Find(model.SettingsId);
            if(setting != null)
            {
                setting.AssumedAllocation = model.AssumedAllocation;
                await _globalview.SaveChangesAsync();
            }
            return Ok();
        }
 

        [HttpGet, Route("api/getallcostcenters/{id}/{sid}")]
        public  IHttpActionResult GetAllCostCenters(int id, long sid)
        {
            var scheduleCostCenter = _globalview.ScheduleCostCenters.Where(x => x.ScheduleID == sid).ToList();
            List<ServiceCostCenterViewModel> objCostCenters = _costAllocationService.GetScheduleCostCenters(sid, scheduleCostCenter).ToList();

            IEnumerable<SccMeterGroupSumsViewModel> MeterGroupSums = objCostCenters
                                                    .SelectMany(x => x.MeterGroups)
                                                    .GroupBy(i => new { id = i.ContractMeterGroupID, name = i.MeterGroupDesc })
                                                    .Select(i => new SccMeterGroupSumsViewModel
                                                    {
                                                        MeterGroupID = i.Key.id,
                                                        Name = i.Key.name,
                                                        Total = i.Sum(t => t.Volume)
                                                    }).ToList();


            return Ok(new { costcenters = objCostCenters, metergroups = MeterGroupSums.OrderBy(x=> x.Name) });
        }
        [HttpPost, Route("api/modifyallcostcenters")]
        public   IHttpActionResult  ModifyAllCostCenters(List<ServiceCostCenterViewModel> models)
        {
           string isValidMsg = string.Empty;
            if (models.Any())
            {
                ExcelRevisionExport ere = new ExcelRevisionExport();

                var scheduleID = models.FirstOrDefault().ScheduleID;
                var _scheduleCostCenters = ere.ServiceCostCenterModelToList(models);


                IEnumerable<SccMeterGroupSumsViewModel> MeterGroupSums = _scheduleCostCenters.GroupBy(l => new { l.MeterGroupID, l.MeterGroupDesc }).Select(x => new SccMeterGroupSumsViewModel
                {
                    MeterGroupID = x.Key.MeterGroupID.Value,
                    Name = x.Key.MeterGroupDesc,
                    Total = x.Sum(t => t.Volume)
                }).ToList();

              isValidMsg =   ere.ValidateMeterGroupVolume(MeterGroupSums, scheduleID);
                if (isValidMsg != "Validated")
                {
                    return Ok(new { message = isValidMsg, costcenters = models });
                }
                else
                {


                    foreach (var model in models)
                    {
                        var scheduleCostCenter = _globalview.ScheduleCostCenters.Where(r => r.CustomerID == model.CustomerID && r.ScheduleID == model.ScheduleID && r.CostCenter == model.CostCenter).ToList();
                        if (scheduleCostCenter.Count > 0)
                        {
                           
                            foreach (var metergroup in model.MeterGroups)
                            {
                                var CostCenter = scheduleCostCenter.Where(x => x.MeterGroupID == metergroup.ContractMeterGroupID).FirstOrDefault();
                                if (CostCenter != null)
                                {
                                    CostCenter.Volume = metergroup.Volume;
                                }
                                else
                                {
                                    var isExists = _globalview.ScheduleCostCenters.Where(r => r.CustomerID == model.CustomerID && r.ScheduleID == model.ScheduleID && r.CostCenter == model.CostCenter && r.MeterGroupID == metergroup.ContractMeterGroupID).FirstOrDefault();
                                    if (metergroup.Volume != 0 && isExists == null)
                                    {
                                        ScheduleCostCenter scc = new ScheduleCostCenter();
                                        scc.ScheduleID = model.ScheduleID;
                                        scc.CustomerID = model.CustomerID;
                                        scc.CostCenter = model.CostCenter;
                                        scc.MeterGroupID = metergroup.ContractMeterGroupID;
                                        scc.Volume = metergroup.Volume;
                                        _globalview.ScheduleCostCenters.Add(scc);
                                    }
                                }
                            }
                            _globalview.SaveChanges();
                        }
                        else
                        {
                            foreach (var metergroup in model.MeterGroups)
                            {
                                if (metergroup.Volume != null)
                                {
                                    ScheduleCostCenter scc = new ScheduleCostCenter();
                                    scc.ScheduleID = model.ScheduleID;
                                    scc.CustomerID = model.CustomerID;
                                    scc.CostCenter = model.CostCenter;
                                    scc.MeterGroupID = metergroup.ContractMeterGroupID;
                                    scc.Volume = metergroup.Volume;
                                    _globalview.ScheduleCostCenters.Add(scc);
                                }

                            }
                            _globalview.SaveChanges();
                        }

                    }

                }
 
            }
            return Ok(new { message = isValidMsg, costcenters = models });
        }
        [HttpGet, Route("api/getallcostcenterschedule/{id}")]
        public async Task<IHttpActionResult> GetSchedulelCostCenters(int id)
        {
            var devices = _globalview.ScheduleDevices.Where(x => x.ScheduleId == id).Select(x => x.EquipmentID).ToList();
            var costcenters = await _globalview.InvoicedEquipmentHistories.Where(x => devices.Contains(x.EquipmentID)).Select(x => x.CostCenter).Distinct().ToListAsync();
            return Ok(new { costcenters = costcenters });
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _globalview.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool CostAllocationSettingExists(int id)
        {
            return _globalview.CostAllocationSettings.Count(e => e.SettingsId == id) > 0;
        }

       
    }
}
 