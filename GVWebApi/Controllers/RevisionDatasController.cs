using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Web.Http;
using System.Web.Http.Description;
using GVWebapi.RemoteData;
using GVWebapi.Helpers;

namespace GVWebapi.Controllers
{
    public class RevisionDatasController : ApiController
    {
        private readonly RevisionDataEntities _revisionDataEntities = new RevisionDataEntities();
        private readonly CoFreedomEntities _coFreedomEntities = new CoFreedomEntities();
        private readonly RevisionDBContext  _revisionDbContext = new RevisionDBContext();
    
        [HttpGet,Route("api/revisiondatas/clientscontract")]
        public IHttpActionResult ClientsContract()
        {
            var clients = _coFreedomEntities.vw_ClientsOnContract.OrderBy(c => c.CustomerName).AsEnumerable();
            return Json(clients);
        }
        
        [HttpGet,Route("api/revisiondatas/getcontracts/{customerid}")]
        public IHttpActionResult GetContracts(int customerId)
        {
            var contracts = _coFreedomEntities.vw_csContractList.OrderBy(c => c.ContractID).Where(c => c.CustomerID == customerId).AsEnumerable();
         
            return Json(contracts);
        }
        
        [HttpGet,Route("api/revisiondatas/getrevisionhistory/{ContractId}")]
        public IHttpActionResult GetRevisionHistory(int contractId)
        {
           
            using (var _dbAudit = new RevisionDataEntities())
            {
                var meterGroups = (from mg in _dbAudit.MeterGroups
                                   where mg.ERPContractID == contractId
                                   select mg).ToList();

                
                 
                
                if (meterGroups.Count > 0)
                {
              
                    IEnumerable<PeriodHistoryWithNote> periods = _dbAudit.PeriodHistoryWithNotes.OrderByDescending(p => p.PeriodDate).Where(p => p.ErpContractID == contractId).ToList();
                    List<RevisionHistoryModel> RevisionModel = new List<RevisionHistoryModel>(); 
                    foreach (var period in periods)
                    {
                        RevisionHistoryModel model = new RevisionHistoryModel();

                        model.peroid = period.PeriodDate;
                        model.detail = _revisionDbContext.PeriodHistoryView.Where(r => r.PeriodDate == period.PeriodDate && r.ERPContractID == contractId).ToList();
                      
                        RevisionModel.Add(model);
                    }
                    
                    return Json(RevisionModel);
                }
                else
                {
                    return Json(HttpStatusCode.NoContent);
                }
            }
                   
        }

        [HttpGet, Route("api/revisiondatas/getreconciliationinvoiced/{ContractId}/{StartDate}/{EndDate}")]
        public IHttpActionResult GetReconciliationInvoiced(int contractId,DateTime startDate,DateTime endDate)
        {

            using (var _dbAudit = new RevisionDataEntities())
            {
                var meterGroups = (from mg in _dbAudit.MeterGroups
                                   where mg.ERPContractID == contractId
                                   select mg).ToList();




                if (meterGroups.Count > 0)
                {
                     RevisionHistoryModel model = new RevisionHistoryModel();
 
                        model.detail = _revisionDbContext.PeriodHistoryView.Where(r => (r.PeriodDate >= startDate && r.PeriodDate <= endDate) && r.ERPContractID == contractId).ToList();

                     

                    return Json(model);
                }
                else
                {
                    return Json(HttpStatusCode.NoContent);
                }
            }

        }


        [HttpGet,Route("api/revisiondatas/getrevisionsummary/{ContractId}")]
        public IHttpActionResult GetRevisionSummary(int contractId)
        {
         
        var Revision = new ExcelRevisionExport();
        var results = Revision.bindRevionDetailSummary(contractId);
         
           return Json(results);
         
        }
        [HttpPost, Route("api/revisiondatas/updatemetergroup/")]
        public IHttpActionResult UpdateMeterGroup(MeterGroup model)
        {
            var meterGroup = _revisionDataEntities.MeterGroups.Find(model.MeterGroupID);
            if(meterGroup != null)
            {
                meterGroup.MeterGroupDesc = model.MeterGroupDesc;
                meterGroup.ERPMeterGroupID = model.ERPMeterGroupID;
                meterGroup.CPP = model.CPP;
                meterGroup.rollovers = model.rollovers;
                _revisionDataEntities.SaveChanges();
            }
            return Ok();
        }

        [HttpPost, Route("api/revisiondatas/addexpense/")]
        public IHttpActionResult AddExpense(BaseHistory model)
        {

            BaseHistory expense = new BaseHistory();
            expense.ContractID = model.ContractID;
            expense.FprBase = model.FprBase;
            expense.PreBase = model.PreBase;
            expense.OverrideDate = model.OverrideDate;
            expense.EndDate = model.EndDate;
            expense.UpdatedOn = DateTime.Now;
            expense.UpdatedBy = model.UpdatedBy;
            expense.Comment = model.Comment;
            _revisionDataEntities.BaseHistories.Add(expense);

            _revisionDataEntities.SaveChanges();
           
            return Ok();
        }
        [HttpPost, Route("api/revisiondatas/updateexpense/")]
        public IHttpActionResult UpdateExpense(BaseHistory model)
        {

            var expense = _revisionDataEntities.BaseHistories.Find(model.ID);
            if (expense != null)
            {
                expense.FprBase = model.FprBase;
                expense.PreBase = model.PreBase;
                expense.OverrideDate = model.OverrideDate;
                expense.EndDate = model.EndDate;
                expense.UpdatedOn = DateTime.Now;
                expense.UpdatedBy = model.UpdatedBy;
                expense.Comment = model.Comment;
                _revisionDataEntities.SaveChanges();
            }
           
            return Ok();
        }
        [HttpPost, Route("api/revisiondatas/updaterevision/")]
        public IHttpActionResult UpdateRevision(RevisionData model)
        {

            var revision = _revisionDataEntities.RevisionDatas.Find(model.RevisionDataID);
            if (revision != null)
            {
                revision.Rollover = model.Rollover;
                revision.Credits = model.Credits;
               
                _revisionDataEntities.SaveChanges();
            }

            return Ok();
        }
        [HttpGet, Route("api/revisiondatas/deleteexpense/")]
        public IHttpActionResult DeleteExpense(int id)
        {

            var expense = _revisionDataEntities.BaseHistories.Find(id);
            if (expense != null)
            {
                _revisionDataEntities.BaseHistories.Remove(expense);
                _revisionDataEntities.SaveChanges();
            }

            return Ok();
        }

        [HttpGet,Route("api/revisiondatas/getmetergroups/{contractid}")]
        public IHttpActionResult GetMeterGroups(int contractId)
        {
            var contractMeterGroup = (from mg in _revisionDataEntities.MeterGroups
                                      where mg.ERPContractID == contractId
                                      select mg);
            var contractdetail = (from c in _coFreedomEntities.SCContracts
                                  where c.ContractID == contractId && c.Active == true
                                 select new
                                    {
                                      StartDate = c.StartDate,
                                     BaseBillingCycleID = c.BaseBillingCycleID,
                                     BaseAccrualCycleID = c.BaseAccrualCycleID
                                 }).FirstOrDefault();
            var BillCycle = String.Empty;
            if (contractdetail.BaseBillingCycleID != null)
                    {
                         
                        {
                            BillCycle = (from bc in _coFreedomEntities.SCBillingCycles
                                             where bc.BillingCycleID == contractdetail.BaseAccrualCycleID
                                             select bc.Description).FirstOrDefault();
                             
                        }
                    }
             else BillCycle = "Quarterly";

            var groupcount = (from gc in _coFreedomEntities.vw_invoicedMeterGroups
                              where gc.ContractID == contractId
                              select gc).Count();
            var baseExpense = (from be in _revisionDataEntities.BaseHistories
                               where be.ContractID == contractId
                               select be).ToList();
            return Json(new { metergroups = contractMeterGroup, invoicedMetergroupcount = groupcount,metergroupcount = contractMeterGroup.Count(), contractstart = contractdetail.StartDate,billingcycle = BillCycle, baseExpenses = baseExpense });
        }

        public IQueryable<RevisionData> GetRevisionDatas()
        {
            return _revisionDataEntities.RevisionDatas;
        }

        [ResponseType(typeof(RevisionData))]
        public async Task<IHttpActionResult> GetRevisionData(int id)
        {
            IEnumerable<RevisionData> revisionData = await _revisionDataEntities.RevisionDatas.Where(r => r.ERPContractID == id).ToListAsync();
            if (revisionData == null)
            {
                return NotFound();
            }

            return Ok(revisionData);
        }

        [ResponseType(typeof(void))]
        public async Task<IHttpActionResult> PutRevisionData(long id, RevisionData revisionData)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            if (id != revisionData.RevisionDataID)
            {
                return BadRequest();
            }

            _revisionDataEntities.Entry(revisionData).State = System.Data.Entity.EntityState.Modified;

            try
            {
                await _revisionDataEntities.SaveChangesAsync();
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!RevisionDataExists(id))
                {
                    return NotFound();
                }
                else
                {
                    throw;
                }
            }

            return StatusCode(HttpStatusCode.NoContent);
        }

        [ResponseType(typeof(RevisionData))]
        public async Task<IHttpActionResult> PostRevisionData(RevisionData revisionData)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            _revisionDataEntities.RevisionDatas.Add(revisionData);
            await _revisionDataEntities.SaveChangesAsync();

            return CreatedAtRoute("DefaultApi", new { id = revisionData.RevisionDataID }, revisionData);
        }

        [ResponseType(typeof(RevisionData))]
        public async Task<IHttpActionResult> DeleteRevisionData(long id)
        {
            RevisionData revisionData = await _revisionDataEntities.RevisionDatas.FindAsync(id);
            if (revisionData == null)
            {
                return NotFound();
            }

            _revisionDataEntities.RevisionDatas.Remove(revisionData);
            await _revisionDataEntities.SaveChangesAsync();

            return Ok(revisionData);
        }

        private bool RevisionDataExists(long id)
        {
            return _revisionDataEntities.RevisionDatas.Count(e => e.RevisionDataID == id) > 0;
        }
        
        private int TotelMonthDifference(DateTime dtThis, DateTime dtOther)
        {
            //Int32 intReturn = 0;

            //dtThis = dtThis.Date.AddDays(-(dtThis.Day - 1));
            //dtOther = dtOther.Date.AddDays(-(dtOther.Day - 1));

            //while (dtOther.Date > dtThis.Date)
            //{
            //    intReturn++;
            //    dtThis = dtThis.AddMonths(1);
            //}

            return ((dtThis.Year - dtOther.Year) * 12) + dtThis.Month - dtOther.Month;
        }
    }
}