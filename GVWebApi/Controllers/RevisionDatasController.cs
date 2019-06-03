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
using GVWebapi.Models;
using GVWebapi.Helpers;
using GVWebapi.Models.Reports;
namespace GVWebapi.Controllers
{
    public class RevisionDatasController : ApiController
    {
       
        private readonly CoFreedomEntities _coFreedomEntities = new CoFreedomEntities();
        private readonly GlobalViewEntities _globalView = new GlobalViewEntities();
    
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
            ExcelRevisionExport revision = new ExcelRevisionExport();
            var RevisionModel = revision.GetRevisionHistory(contractId);

             if(RevisionModel != null)
            { 
                return Json(RevisionModel);
            }
            else
            {
                return Json(HttpStatusCode.NoContent);
            }
               
        }


        [HttpGet, Route("api/revisiondatas/chart/{customerId}")]
        public IHttpActionResult GetContractVSActualHistory(int customerId)
        {
          
            var contract = _coFreedomEntities.vw_csContractList.OrderBy(c => c.ContractID).Where(c => c.CustomerID == customerId).FirstOrDefault();
            ExcelRevisionExport revision = new ExcelRevisionExport();
            var revs = revision.GetRevisionHistory(contract.ContractID).Reverse();
            var mgs = _globalView.RevisionMeterGroups.Where(x => x.ERPContractID == contract.ContractID);

           
            var chardata = new List<ContractedVolumeReport>();
           
            foreach(var rev in revs)
            {
                foreach(var detail in rev.detail)
                {
                    var revisiondata = new ContractedVolumeReport();
                    revisiondata.Period = rev.peroid.Value.ToString("MMM, yyyy");
                    revisiondata.MeterGroup = detail.ContractMeterGroupID.Value;
                    revisiondata.MeterGroupDesc = detail.MeterGroup;
                    revisiondata.ContractedVolume = detail.ContractVolume;
                    revisiondata.ActualVolume = detail.ActualVolume;
                    chardata.Add(revisiondata);
                }
               
            }
            var chartlist = new List<dynamic>();
           
            var random = new Random();
            var i = 1;
            foreach (var mg in mgs)
            {
                var chardata2 = chardata.Where(x => x.MeterGroup == mg.ERPMeterGroupID);
                var MeterGroup = chardata2.Select(x => x.MeterGroupDesc).First();
                var color = String.Format("#{0:X6}", random.Next(0x1000000));
                var a = new { idx = i++, show = false, label = "Contracted " + MeterGroup, color = color, data = chardata2.Select(x => new object[] { x.Period, x.ContractedVolume }) };
                chartlist.Add(a);
                var b = new { idx = i++, show = false, label = "Actual " + MeterGroup, color = color + "82", data = chardata2.Select(x => new object[] { x.Period, x.ActualVolume }) };
                chartlist.Add(b);
              

            }
            

            return Json(chartlist);

        }


        [HttpGet, Route("api/revisiondatas/volumehistorychart/{customerId}")]
        public IHttpActionResult GetContractVSActualHistory2(int customerId)
        {
            var contract = _coFreedomEntities.vw_csContractList.OrderBy(c => c.ContractID).Where(c => c.CustomerID == customerId).FirstOrDefault();
            var ContractDate = DateTime.Now.AddMonths(-48);
            ExcelRevisionExport revision = new ExcelRevisionExport();
            var revs = revision.GetRevisionHistory(contract.ContractID).Where(x => x.peroid >= ContractDate).Reverse();
            var mgs = _globalView.RevisionDataViews.Where(r => r.ContractID == contract.ContractID && r.OverageToDate >= ContractDate).Select(x => new { x.ContractMeterGroupID, x.MeterGroup }).ToList().Distinct();
            var chartdata = new List<dynamic>();
            var periodDate = revs.First().peroid;
            foreach (var rev in revs)
            {
                var datediff = (rev.peroid - periodDate).Value.Days;
                periodDate = rev.peroid.Value;
                if(datediff > 90)
                {
                    var loop = datediff / 90;
                   for( var i = 1; loop < i; i++)
                    {


                    }

                }

                var chartItem = new Dictionary<string, string>();
                chartItem.Add("Period", rev.peroid.Value.ToString("MMM, yyyy"));
                foreach (var detail in rev.detail)
                {


                    chartItem.Add(detail.MeterGroup + " Contracted", detail.ContractVolume.Value.ToString());
                    chartItem.Add(detail.MeterGroup + " Actual", detail.ActualVolume.Value.ToString());
                    
                }
                chartdata.Add(chartItem);
            }
           
            var seriesdata = new List<dynamic>();
            foreach (var mg in mgs)
            {
                var seriesItem = new Dictionary<string, string>();
                seriesItem.Add("dataField", mg.MeterGroup + " Contracted");
                seriesdata.Add(seriesItem);
                var seriesItem2 = new Dictionary<string, string>();
                seriesItem2.Add("dataField", mg.MeterGroup + " Actual");
                seriesdata.Add(seriesItem2);
            }


            return Json( new { data = chartdata, series = seriesdata });

        }

        [HttpGet, Route("api/reports/volumehistorychart/{customerId}")]
        public IHttpActionResult ContractVSActualHistoryChart(int customerId)
        {
            var contract = _coFreedomEntities.vw_csContractList.OrderBy(c => c.ContractID).Where(c => c.CustomerID == customerId).FirstOrDefault();
            var ContractDate = DateTime.Now.AddMonths(-48);
            ExcelRevisionExport revision = new ExcelRevisionExport();
            var revs = revision.GetRevisionHistory(contract.ContractID).Where(x => x.peroid >= ContractDate).Reverse();
            var mgs = _globalView.RevisionDataViews.Where(r => r.ContractID == contract.ContractID && r.OverageToDate >= ContractDate).Select(x => new { x.ContractMeterGroupID, x.MeterGroup }).ToList().Distinct();
            var chartdata = new List<dynamic>();
            var periodDate = revs.First().peroid;
            foreach (var rev in revs)
            {
                var datediff = (rev.peroid - periodDate).Value.Days;
                periodDate = rev.peroid.Value;

                if (datediff > 99)
                {

                    var loop = datediff / 90;

                    for (var i = 1; i <= loop; i++)
                    {
                        var chartItemx = new Dictionary<string, string>();
                        if (i == loop)
                        {
                            chartItemx.Add("Period", rev.peroid.Value.ToString("MMM, yyyy"));
                        }
                        else
                        {
                            chartItemx.Add("Period", rev.peroid.Value.AddDays(-90).ToString("MMM, yyyy"));
                        }

                        foreach (var detail in rev.detail)
                        {
                            var contractedVol = detail.ContractVolume.Value / loop;
                            var actualVol = detail.ActualVolume.Value / loop;

                            chartItemx.Add(detail.MeterGroup + " Contracted", contractedVol.ToString());
                            chartItemx.Add(detail.MeterGroup + " Actual", actualVol.ToString());

                        }
                        chartdata.Add(chartItemx);
                    }

                }
                else
                {
                    var chartItem = new Dictionary<string, string>();
                    chartItem.Add("Period", rev.peroid.Value.ToString("MMM, yyyy"));
                    foreach (var detail in rev.detail)
                    {


                        chartItem.Add(detail.MeterGroup + " Contracted", detail.ContractVolume.Value.ToString());
                        chartItem.Add(detail.MeterGroup + " Actual", detail.ActualVolume.Value.ToString());

                    }
                    chartdata.Add(chartItem);
                }


            }

            var c_seriesdata = new List<dynamic>();
            var a_seriesdata = new List<dynamic>();
            foreach (var mg in mgs)
            {
                var seriesItem = new Dictionary<string, string>();
                seriesItem.Add("dataField", mg.MeterGroup + " Contracted");
                seriesItem.Add("opacity", "1.0");
                seriesItem.Add("lineWidth", "4");
                seriesItem.Add("dashStyle", "4,4");
                c_seriesdata.Add(seriesItem);
                var seriesItem2 = new Dictionary<string, string>();
                seriesItem2.Add("dataField", mg.MeterGroup + " Actual");
                seriesItem2.Add("opacity", "0.4");
                a_seriesdata.Add(seriesItem2);

            }


            return Json(new { data = chartdata, cseries = c_seriesdata, aseries = a_seriesdata });

        }


        [HttpGet, Route("api/revisiondatas/getreconciliationinvoiced/{ContractId}/{StartDate}/{EndDate}")]
        public IHttpActionResult GetReconciliationInvoiced(int contractId,DateTime startDate,DateTime endDate)
        {

        var periods =  _coFreedomEntities.vw_REVisionInvoices.Where(r => (r.PeriodDate >= startDate && r.PeriodDate <= endDate) && r.ContractID == contractId).ToList();

            List<RevisionHistoryModel> RevisionModel = new List<RevisionHistoryModel>();
            foreach (var period in periods)
            {
                RevisionHistoryModel model = new RevisionHistoryModel(); 
               // model.detail = ExcelRevisionExport.GetRevisionHistory(contractId);
                RevisionModel.Add(model);
            }
            return Json(RevisionModel);
        }


        [HttpGet,Route("api/revisiondatas/getrevisionsummary/{ContractId}")]
        public IHttpActionResult GetRevisionSummary(int contractId)
        {
         
        var Revision = new ExcelRevisionExport();
        var results = Revision.RevisionSummary(contractId);
         
           return Json(results);
         
        }
        [HttpPost, Route("api/revisiondatas/updatemetergroup/")]
        public IHttpActionResult UpdateMeterGroup(RevisionMeterGroup model)
        {
            var meterGroup = _globalView.RevisionMeterGroups.Find(model.MeterGroupID);
            if(meterGroup != null)
            {
                meterGroup.MeterGroupDesc = model.MeterGroupDesc;
                meterGroup.ERPMeterGroupID = model.ERPMeterGroupID;
                meterGroup.CPP = model.CPP;
                meterGroup.rollovers = model.rollovers;
                _globalView.SaveChanges();
            }
            return Ok();
        }
         

        [HttpPost, Route("api/revisiondatas/addexpense/")]
        public IHttpActionResult AddExpense(RevisionBaseExpense model)
        {

            RevisionBaseExpense expense = new RevisionBaseExpense();
            expense.ContractID = model.ContractID;
            expense.FprBase = model.FprBase;
            expense.PreBase = model.PreBase;
            expense.OverrideDate = model.OverrideDate;
            expense.EndDate = model.EndDate;
            expense.UpdatedOn = DateTime.Now;
            expense.UpdatedBy = model.UpdatedBy;
            expense.Comment = model.Comment;
            _globalView.RevisionBaseExpenses.Add(expense);

            _globalView.SaveChanges();
           
            return Ok();
        }
        [HttpPost, Route("api/revisiondatas/updateexpense/")]
        public IHttpActionResult UpdateExpense(RevisionBaseExpense model)
        {

            var expense = _globalView.RevisionBaseExpenses.Find(model.ID);
            if (expense != null)
            {
                expense.FprBase = model.FprBase;
                expense.PreBase = model.PreBase;
                expense.OverrideDate = model.OverrideDate;
                expense.EndDate = model.EndDate;
                expense.UpdatedOn = DateTime.Now;
                expense.UpdatedBy = model.UpdatedBy;
                expense.Comment = model.Comment;
                _globalView.SaveChanges();
            }
           
            return Ok();
        }
        [HttpPost, Route("api/revisiondatas/updaterevision/")]
        public IHttpActionResult UpdateRevision(RevisionDataModel model)
        {

            var revision = _globalView.RevisionDatas.Find(model.RevisionID);
            if (revision != null)
            {
                revision.Rollover = model.Rollover;
                revision.Credits = model.CreditAmount;

                _globalView.SaveChanges();
            } else
            {
                RevisionData revisionData = new RevisionData();
                revisionData.ContractID = model.ContractID;
                revisionData.InvoiceID = model.InvoiceID;
                revisionData.MeterGroupID = model.ContractMeterGroupID.Value;
                revisionData.Rollover = model.Rollover;
                revisionData.Credits = model.CreditAmount;
                _globalView.RevisionDatas.Add(revisionData);
                _globalView.SaveChanges();
            }

            return Ok();
        }
        [HttpGet, Route("api/revisiondatas/deleteexpense/")]
        public IHttpActionResult DeleteExpense(int id)
        {

            var expense = _globalView.RevisionBaseExpenses.Find(id);
            if (expense != null)
            {
                _globalView.RevisionBaseExpenses.Remove(expense);
                _globalView.SaveChanges();
            }

            return Ok();
        }

        [HttpGet,Route("api/revisiondatas/getmetergroups/{contractid}")]
        public IHttpActionResult GetMeterGroups(int contractId)
        {
            var contractMeterGroup = (from mg in _globalView.RevisionMeterGroups
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
            var baseExpense = (from be in _globalView.RevisionBaseExpenses
                               where be.ContractID == contractId
                               select be).OrderByDescending(o=> o.OverrideDate).ToList();
            return Json(new { metergroups = contractMeterGroup, invoicedMetergroupcount = groupcount,metergroupcount = contractMeterGroup.Count(), contractstart = contractdetail.StartDate,billingcycle = BillCycle, baseExpenses = baseExpense });
        }
 

        [HttpGet, Route("api/revisiondatas/getrevisiondatas/{contractid}")]
        public IHttpActionResult AllRevisionDatas(int contractId)
        {
            var contractMeterGroup = (from mg in _globalView.RevisionMeterGroups
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
            var baseExpense = (from be in _globalView.RevisionBaseExpenses
                               where be.ContractID == contractId
                               select be).OrderByDescending(o => o.OverrideDate).ToList();

            var Revision = new ExcelRevisionExport();
            var summary = Revision.RevisionSummary(contractId);
            var history = Revision.GetRevisionHistory(contractId);
            return Json(new { metergroups = contractMeterGroup, history = history, summary = summary, invoicedMetergroupcount = groupcount, metergroupcount = contractMeterGroup.Count(), contractstart = contractdetail.StartDate, billingcycle = BillCycle, baseExpenses = baseExpense });
        }

        public IQueryable<vw_RevisionInvoiceHistory> GetRevisionDatas()
        {
            return _coFreedomEntities.vw_RevisionInvoiceHistory;
        }

        [ResponseType(typeof(RevisionData))]
        public async Task<IHttpActionResult> GetRevisionData(int id)
        {
            IEnumerable<vw_RevisionInvoiceHistory> revisionData = await _coFreedomEntities.vw_RevisionInvoiceHistory.Where(r => r.ContractID == id).ToListAsync();
            if (revisionData == null)
            {
                return NotFound();
            }

            return Ok(revisionData);
        }

         
         
        
        private int TotalMonthDifference(DateTime dtThis, DateTime dtOther)
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