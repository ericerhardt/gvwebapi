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
        
        [HttpGet,Route("api/revisiondatas/getrevisionsummary/{ContractId}")]
        public IHttpActionResult GetRevisionSummary(int contractId)
        {
            using (var _dbAudit = new RevisionDataEntities())
            {
                var VisionDataList = new List<VisionData>();

                var revisions = (from r in _dbAudit.RevisionDataViews
                                 where r.ERPContractID == contractId
                                 orderby r.PeriodDate descending
                                 select new
                                 {
                                     ERPContarctID = contractId,
                                     PeriodDates = r.PeriodDate.Value,
                                     StartDate = r.StartDate.Value,
                                     Difference = ((r.PeriodDate.Value.Year - r.StartDate.Value.Year) * 12) + (r.PeriodDate.Value.Month - r.StartDate.Value.Month) + 1,
                                     PreFPRBaseExpense = r.ClientContractBase,
                                     PreFPROverageExpense = r.OverageNoFPR,
                                     FPRBaseExpense = r.FBRContractBase,
                                     FPROverageExpense = r.OverageExpense,
                                     r.Credits,
                                     PeriodSavings = (r.OverageNoFPR + r.ClientContractBase) - (r.OverageExpense + r.FBRContractBase),
                                     PctSavings = r.NetSavings.Value != 0 && r.OverageNoFPR.Value != 0 ? r.NetSavings / (r.ClientContractBase + r.OverageNoFPR) : 0
                                 }).ToList();
                var summary = (from s in revisions
                               group s by new
                               {

                                   clientPeriodDates = s.PeriodDates,
                                   clientPeriodStart = s.StartDate,
                                   diff = s.Difference,
                                   fprCost = s.FPRBaseExpense,
                                   clientCost = s.PreFPRBaseExpense

                               }
                                   into v
                               select new
                               {
                                   v.Key.clientPeriodDates,
                                   clientPeriodDate = v.Key.clientPeriodDates,
                                   clientStartDate = v.Key.clientPeriodStart,
                                   fprOverageCost = v.Sum(o => o.FPROverageExpense),
                                   prefprOverageCost = v.Sum(o => o.PreFPROverageExpense),
                                   fprCost = v.Key.fprCost * v.Key.diff,
                                   clientOverageCost = v.Sum(o => o.PreFPROverageExpense),
                                   clientCost = v.Key.clientCost * v.Key.diff,
                                   credits = v.Sum(o => o.Credits),
                                   savings = v.Sum(o => o.PeriodSavings),
                                   pct = v.Sum(o => o.PctSavings)
                               }).Distinct();


                foreach (var revision in summary)
                {
                    Int32 monthsDiff = TotelMonthDifference(revision.clientPeriodDate, revision.clientStartDate);
                    var a = new VisionData();
                    a.ERPContractID = contractId;
                  //  a.clientPeriodDates = revision.clientPeriodDates.AddMonths(-monthsDiff).ToString("MMM") + " - " + revision.clientPeriodDates.ToString("MMM yyyy");
                    a.clientPeriodDates = revision.clientPeriodDates.AddMonths(-monthsDiff).ToString("MMM") + " - " + revision.clientPeriodDates.ToString("MMM yyyy"); 
                    a.fprOverageCost = Convert.ToDouble(revision.fprOverageCost);
                    a.clientOverageCost = Convert.ToDouble(revision.prefprOverageCost);
                    a.fprCost = Convert.ToDouble(revision.fprCost);
                    a.clientCost = Convert.ToDouble(revision.clientCost);
                    a.credits = Convert.ToDouble(revision.credits);
                    a.savings = Convert.ToDouble(revision.savings);
                    a.pct = Convert.ToDouble(revision.pct);
                    VisionDataList.Add(a);
                }
                try
                {
                    var summary2 = (from s in VisionDataList
                                    group s by new
                                    {

                                        clientPeriodDates = s.clientPeriodDates,
                                        clientPeriodStart = s.clientPeriodDates,
                                        fprCost = s.fprCost,
                                        clientCost = s.clientCost

                                    }
                                        into v
                                    select new
                                    {
                                        clientPeriodDates = v.Key.clientPeriodDates,
                                        clientPeriodDate = v.Key.clientPeriodDates,
                                        clientStartDate = v.Key.clientPeriodStart,
                                        fprOverageCost = v.Sum(o => o.fprOverageCost),
                                        prefprOverageCost = v.Sum(o => o.clientOverageCost),
                                        fprCost = v.Key.fprCost,
                                        clientOverageCost = v.Sum(o => o.clientOverageCost),
                                        clientCost = v.Key.clientCost,
                                        credits = v.Sum(o => o.credits),
                                        savings = v.Sum(o => o.savings),
                                        pct = v.Sum(o => o.pct)
                                    }).Distinct();

                }
                catch (Exception e)
                {
                    throw(e);
                }

              var results = VisionDataList.OrderByDescending(o => o.clientPeriodDates);
                return Json(results);
                 
            }

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
            return Json(new { metergroups = contractMeterGroup, invoicedMetergroupcount = groupcount,metergroupcount = contractMeterGroup.Count(), contractstart = contractdetail.StartDate,billingcycle = BillCycle });
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