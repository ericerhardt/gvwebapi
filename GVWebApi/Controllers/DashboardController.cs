using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GVWebapi.RemoteData;
using GVWebapi.Models;
using GVWebapi.Helpers;

namespace GVWebapi.Controllers
{
    public class DashboardController : ApiController
    {
        private readonly CoFreedomEntities _context = new CoFreedomEntities();
        private readonly GlobalViewEntities _db = new GlobalViewEntities();
        private CustomerPortalEntities db = new CustomerPortalEntities();
        private RevisionDataEntities rev = new RevisionDataEntities();
        [HttpGet]
        [Route("api/savingbycategory/{ClientID}")]
        public IHttpActionResult SavingsByCategory(int ClientID)
        {
            var contract = _context.vw_csContractList.OrderBy(c => c.ContractID).Where(c => c.CustomerID == ClientID).First().ContractID;
          var Replacements = _db.AssetReplacements.Where(d => d.CustomerID == ClientID).Sum(c => c.ReplacementValue).ToString() ?? "0";
           var CostAvoidance =  db.CostAvoidances.Where(c => c.CustomerID == ClientID).Sum(c => c.TotalSavingsCost).ToString() ?? "0";
           var RollOvers = rev.RolloverViews.Where(c => c.CustomerID == ClientID).Sum(r => r.Rollover * r.CPP).ToString() ??"0";
            var Revision = visionSummary(contract).ToString() ?? "0" ;

        
            var ret = new[]
           {
                   new { label="Replacements", color = "#4acab4", data =   Replacements },
                   new { label="REVision", color = "#ffea88" , data =  Revision},
                   new { label="Cost Avoidance", color = "#ff8153" , data =  CostAvoidance},
                   new { label="Rollovers", color = "#878bb6" , data =  RollOvers}
            };
            return Json(ret);
        }
        [HttpGet]
        [Route("api/maplocations/{ClientID}")]
        public IHttpActionResult GetMapLocations(int ClientID)
        {
          
          var maplocation =   _context.vw_maplocations.Where(c => c.LocationID == ClientID && c.Active == true).Select(c => new { location = c.Latitude +", " + c.Longitude, name = c.CustomerName }).ToList().Distinct();
            return Json(maplocation);
        }
        // GET: api/Dashboard/5
        public string Get(int id)
        {
            return "value";
        }

        // POST: api/Dashboard
       

        private double? visionSummary(int ContractID)
        {
            using (var _dbAudit = new RevisionDataEntities())
            {
                var VisionDataList = new List<VisionData>();

                var revisions = (from r in _dbAudit.RevisionDataViews
                                 where r.ERPContractID == ContractID
                                 orderby r.PeriodDate descending
                                 select new
                                 {
                                     ERPContarctID = ContractID,
                                     PeriodDates = r.PeriodDate.Value,
                                     StartDate = r.StartDate.Value,
                                     Difference = ((r.PeriodDate.Value.Year - r.StartDate.Value.Year) * 12) + (r.PeriodDate.Value.Month - r.StartDate.Value.Month) + 1,
                                     PreFPRBaseExpense = r.ClientContractBase,
                                     PreFPROverageExpense = r.OverageNoFPR,
                                     FPRBaseExpense = r.FBRContractBase,
                                     FPROverageExpense = r.OverageExpense,
                                     Credits = r.Credits,
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
                                   clientPeriodDates = v.Key.clientPeriodDates,
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
                    a.ERPContractID = ContractID;
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
                                        
                                        fprOverageCost = v.Sum(o => o.fprOverageCost),
                                        prefprOverageCost = v.Sum(o => o.clientOverageCost),
                                        fprCost = v.Key.fprCost,
                                        clientOverageCost = v.Sum(o => o.clientOverageCost),
                                        clientCost = v.Key.clientCost,
                                        credits = v.Sum(o => o.credits),
                                        savings = v.Sum(o => o.savings),
                                        pct = v.Sum(o => o.pct)
                                    }).Distinct();
                    var savings = summary2.Sum(s => s.savings).Value;
                    var credits = summary2.Sum(s => s.credits).Value;
                    return savings + credits;
                }
                catch (Exception e)
                {
                    throw (e);
                }
              
            }
        }
        private Int32 TotelMonthDifference(DateTime dtThis, DateTime dtOther)
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
        [HttpGet]
        [Route("api/weeklycalls/{CustomerID}")]
        public IHttpActionResult GetCustomerservicecallCounts(int CustomerID)
        {
            DateTime today = DateTime.Now;
            DateTime ThisWeekStart = today.FirstDayOfWeek();
            DateTime ThisWeekEnd = today.LastDayOfWeek();
            DateTime LastWeekStart = ThisWeekStart.AddDays(-7);
            DateTime LastWeekEnd = ThisWeekEnd.AddDays(-7);
            ServiceCallsViewModel serviceCall = new ServiceCallsViewModel();



            List<ServiceCallsViewModel> serviceCalls = new List<ServiceCallsViewModel>();

            serviceCall.data = _context.vw_CSServiceCallHistory
                                            .Where(c => c.CustomerID == CustomerID && (c.Date <= LastWeekEnd && c.Date >= LastWeekStart && c.v_Status != "Canceled") && !c.Description.Contains("Supply"))
                                            .OrderBy(c=> c.Date).ToList()
                                            .GroupBy(model => model.Date.ToString("ddd"), (i, models) => new WeeklyCallTotals { day = i, totalcalls = models.Count()  });
            var daysOfWeek = Enum.GetValues(typeof(DayOfWeek))
                 .OfType<DayOfWeek>()
                 .Where(day => day > DayOfWeek.Sunday && day < DayOfWeek.Saturday)
                 .OrderBy(day => day < DayOfWeek.Monday);

            List<WeeklyCallTotals> lastweeklycalls = new List<WeeklyCallTotals>();
            List<WeeklyCallTotals> thisweeklycalls = new List<WeeklyCallTotals>();
            foreach (var day in daysOfWeek)
            {
                    string daymod = day.ToString().Substring(0, 3);
                    var data = serviceCall.data.Where(x => x.day == daymod).FirstOrDefault();
                    if (data!= null)
                    {
                    lastweeklycalls.Add(new WeeklyCallTotals { day = day.ToString().Substring(0, 3), totalcalls = data.totalcalls });
                    }
                    else
                    {
                    lastweeklycalls.Add(new WeeklyCallTotals { day = day.ToString().Substring(0, 3), totalcalls = 0 });
                    }
                
            }

            ServiceCallsViewModel serviceCall2 = new ServiceCallsViewModel();

            serviceCall2.data = _context.vw_CSServiceCallHistory.Where(c => c.CustomerID == CustomerID && (c.Date <= ThisWeekEnd && c.Date >= ThisWeekStart && c.v_Status != "Canceled") && !c.Description.Contains("Supply")).OrderBy(c => c.Date).ToList().GroupBy(model => model.Date.ToString("ddd"), (i, models) => new WeeklyCallTotals { day = i, totalcalls = models.Count() });

            foreach (var day in daysOfWeek)
            {
                string daymod = day.ToString().Substring(0, 3);
                var data = serviceCall2.data.Where(x => x.day == daymod).FirstOrDefault();
                if (data != null)
                {
                    thisweeklycalls.Add(new WeeklyCallTotals { day = day.ToString().Substring(0, 3), totalcalls = data.totalcalls });
                }
                else
                {
                    thisweeklycalls.Add(new WeeklyCallTotals { day = day.ToString().Substring(0, 3), totalcalls = 0 });
                }

            }

            var ret = new[]
            {
                  new { label="Last Week", color = "#aad874", data = lastweeklycalls.Select(x=>new object[]{ x.day,  x.totalcalls  })},
                  new { label="This Week", color = "#7dc7df" , data = thisweeklycalls.Select(x=> new object[]{ x.day, x.totalcalls  })}
            };

            return Json(ret);
        }
    }
}
