using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;
using GVWebapi.Models;
using GVWebapi.Helpers;
using GVWebapi.Models.Schedules;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace GVWebapi.Controllers
{
    public class CostAllocationController : ApiController
    {
        private GlobalViewEntities db = new GlobalViewEntities();
        private ExcelRevisionExport er = new ExcelRevisionExport();
        // GET: api/CostAllocation/5
        [HttpGet,Route("api/allocationsettings/{id}")]
        public async Task<IHttpActionResult> GetCostAllocationSetting(int id)
        {
            CostAllocationSetting costAllocationSetting = new CostAllocationSetting();
            costAllocationSetting = await db.CostAllocationSettings.Where(x =>x.CustomerID == id).FirstOrDefaultAsync();
            List<CostAllocationSettingsViewModel> MeterGroupsModel = new List<CostAllocationSettingsViewModel>();
            if (costAllocationSetting == null)
            {
                CostAllocationSetting row = new CostAllocationSetting();
                row.CustomerID = id;
                row.AssumedAllocation = true;
                db.CostAllocationSettings.Add(row);
                db.SaveChanges();
                costAllocationSetting = await db.CostAllocationSettings.Where(x => x.CustomerID == id).FirstOrDefaultAsync();
            }
          
            List<CostAllocationMeterGroup> AllocationMetergroups = new List<CostAllocationMeterGroup>();
            AllocationMetergroups = await db.CostAllocationMeterGroups.Where(x => x.CustomerID == id).ToListAsync();
            if (AllocationMetergroups.Count() == 0)
            {
                var ContractID = er.GetContractID(id);
                var revMeterGroups = db.RevisionMeterGroups.Where(x => x.ERPContractID == ContractID);
                foreach (var mg in revMeterGroups)
                {

                    CostAllocationMeterGroup metergroup = new CostAllocationMeterGroup()
                    {
                        CustomerID = id,
                        MeterGroupID = (int)mg.MeterGroupID,
                        ExcessCPP = 0.00M
                    };
                    db.CostAllocationMeterGroups.Add(metergroup);
                }
                 db.SaveChanges();
                AllocationMetergroups = await db.CostAllocationMeterGroups.Where(x => x.CustomerID == id).ToListAsync();
            } 
                foreach(var mg in AllocationMetergroups)
                {
                    
                    CostAllocationSettingsViewModel model = new CostAllocationSettingsViewModel()
                    {
                        SettingsMeterGroupID = mg.SettingsMeterGroupID,
                        CustomerID = mg.CustomerID,
                        MeterGroupID = mg.MeterGroupID,
                        MeterGroupDesc = db.RevisionMeterGroups.Find(mg.MeterGroupID).MeterGroupDesc,
                        ExcessCPP = mg.ExcessCPP
                    };
                    MeterGroupsModel.Add(model);
                }
               
            

            return Ok( new { settings = costAllocationSetting, metergorups = MeterGroupsModel });
        }
        [HttpPost, Route("api/saveexcessmetergroups")]
        public async Task<IHttpActionResult> SaveMeterGroupExcessCPP(IEnumerable<CostAllocationSettingsViewModel> models )
        {
            foreach(var model in models)
            {
                var metergroup = db.CostAllocationMeterGroups.Find(model.SettingsMeterGroupID);
                {
                    metergroup.ExcessCPP = model.ExcessCPP;
                };
               
            }
            await db.SaveChangesAsync();
        

            return Ok();
        }

        [HttpPost, Route("api/costallocationsave")]
        public async Task<IHttpActionResult> SaveCostAllocationSave(CostAllocationSetting model)
        {
            
            var setting = db.CostAllocationSettings.Find(model.SettingsId);
            if(setting != null)
            {
                setting.AssumedAllocation = model.AssumedAllocation;
                await db.SaveChangesAsync();
            }
            return Ok();
        }
        [HttpGet, Route("api/getallcostcenters/{id}/{sid}")]
        public async Task<IHttpActionResult> GetAllCostCenters(int id, long sid)
        {
            var schedule = db.Schedules.Where(x => x.ScheduleId == sid).FirstOrDefault();
            var costcenters = await db.InvoicedEquipmentHistories
                                      .Where(x => x.CustomerID == id && x.CostCenter != "" &&  x.PeriodDate <= schedule.ExpiredDateTime)
                                      .Select(x => x.CostCenter).Distinct().ToListAsync();
            var metergroups = await db.CostAllocationMeterGroups.Where(x => x.CustomerID == id).Select(x => x.MeterGroupID).Distinct().ToListAsync();

            List<ServiceCostCenterViewModel> objCostCenters = new List<ServiceCostCenterViewModel>();
            foreach (var costcenter in costcenters)
            {
                ServiceCostCenterViewModel obj = new ServiceCostCenterViewModel();
                obj.Status = 0;
                obj.CostCenter = costcenter;
                obj.MeterGroups = new List<MeterGroupCostCenter>();
                foreach (var metergroup in metergroups)
                {
                   
                    MeterGroupCostCenter mg = new MeterGroupCostCenter();
                    mg.MeterGroupDesc = db.RevisionMeterGroups.Find(metergroup).MeterGroupDesc;
                    mg.MeterGroupID = metergroup;
                    mg.Volume = 0;
                    obj.MeterGroups.Add(mg);
                }
                obj.ScheduleID = sid;
                obj.CustomerID = id;
                objCostCenters.Add(obj);
            }

            var scheduleCostCenters = db.ScheduleCostCenters.Where(x => x.ScheduleID == sid).ToList();

            foreach (var objcostcenter in objCostCenters)
            {
                foreach (var metergroup in objcostcenter.MeterGroups)
                {

                    var scheduleServiceMg = scheduleCostCenters.Where(x => x.CostCenter == objcostcenter.CostCenter && x.MeterGroupID == metergroup.MeterGroupID).FirstOrDefault();
                    if (scheduleServiceMg != null)
                    {
                        objcostcenter.ScheduleCostCenterID = scheduleServiceMg.ScheduleCostCenterID;
                        objcostcenter.Status = 1;
                        metergroup.Volume = scheduleServiceMg.Volume;

                    }
                }
            }

            var MeterGroupSums = scheduleCostCenters.GroupBy(l => l.MeterGroupID).Select(x => new
            {
                Name = db.RevisionMeterGroups.Find(x.Key).MeterGroupDesc,
                Total = x.Sum(t => t.Volume)
            });


            return Ok(new { costcenters = objCostCenters, metergroups = MeterGroupSums.OrderBy(x=> x.Name) });
        }
        [HttpPost, Route("api/modifyallcostcenters")]
        public   IHttpActionResult  ModifyAllCostCenters(List<ServiceCostCenterViewModel> models)
        {
            if (models.Any())
            {
                var scheduleID = models.FirstOrDefault().ScheduleID;
                var _scheduleCostCenters = db.ScheduleCostCenters.Where(x => x.ScheduleID == scheduleID).ToList();
                var MeterGroupSums = _scheduleCostCenters.GroupBy(l => l.MeterGroupID).Select(x => new
                {
                    Name = db.RevisionMeterGroups.Find(x.Key).MeterGroupDesc,
                    Total = x.Sum(t => t.Volume)
                });

                foreach (var model in models)
                {
                    var scheduleCostCenters = db.ScheduleCostCenters.Find(model.ScheduleCostCenterID);
                    if (scheduleCostCenters != null)
                    {
                         foreach(var metergroup in model.MeterGroups)
                         {
                            if(scheduleCostCenters.MeterGroupID == metergroup.MeterGroupID)
                            {
                                scheduleCostCenters.Volume = metergroup.Volume;
                            } else
                            {
                                var isExists = db.ScheduleCostCenters.Where(r => r.CustomerID == model.CustomerID && r.ScheduleID == model.ScheduleID && r.CostCenter == model.CostCenter && r.MeterGroupID == metergroup.MeterGroupID).FirstOrDefault();
                                if (metergroup.Volume != 0 && isExists == null)
                                {
                                    ScheduleCostCenter scc = new ScheduleCostCenter();
                                    scc.ScheduleID = model.ScheduleID;
                                    scc.CustomerID = model.CustomerID;
                                    scc.CostCenter = model.CostCenter;
                                    scc.MeterGroupID = metergroup.MeterGroupID;
                                    scc.Volume = metergroup.Volume;
                                    db.ScheduleCostCenters.Add(scc);
                                }
                            }                         
                         }
                         db.SaveChanges();                         
                    } else
                    {
                        foreach (var metergroup in model.MeterGroups)
                        {
                            if (metergroup.Volume != 0) {  
                                ScheduleCostCenter scc = new ScheduleCostCenter();
                                scc.ScheduleID = model.ScheduleID;
                                scc.CustomerID = model.CustomerID;
                                scc.CostCenter = model.CostCenter;
                                scc.MeterGroupID = metergroup.MeterGroupID;
                                scc.Volume = metergroup.Volume;
                                db.ScheduleCostCenters.Add(scc);
                            }

                        }
                        db.SaveChanges();
                    }

                }
                
            }

            return Ok(new { costcenters = models });
        }
        [HttpGet, Route("api/getallcostcenterschedule/{id}")]
        public async Task<IHttpActionResult> GetSchedulelCostCenters(int id)
        {
            var devices = db.Devices.Where(x => x.ScheduleId == id).Select(x => x.EquipmentId).ToList();
            var costcenters = await db.InvoicedEquipmentHistories.Where(x => devices.Contains(x.EquipmentID)).Select(x => x.CostCenter).Distinct().ToListAsync();
            return Ok(new { costcenters = costcenters });
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }

        private bool CostAllocationSettingExists(int id)
        {
            return db.CostAllocationSettings.Count(e => e.SettingsId == id) > 0;
        }

       
    }
}
 