using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GVWebapi.RemoteData;
using GVWebapi.Models;
using GVWebapi.Helpers;
using System.Data.Entity;
 

namespace GVWebapi.Controllers
{
    public class DashboardController : ApiController
    {
        private readonly CoFreedomEntities _context = new CoFreedomEntities();
        private readonly GlobalViewEntities _db = new GlobalViewEntities();
        private CustomerPortalEntities db = new CustomerPortalEntities();
      
        [HttpGet]
        [Route("api/savingbycategory/{ClientID}")]
        public IHttpActionResult SavingsByCategory(int ClientID)
        {
            ExcelRevisionExport er = new ExcelRevisionExport();
            var ReplacementValue = 0.00M;
            var contract = _context.vw_csContractList.OrderBy(c => c.ContractID).Where(c => c.CustomerID == ClientID).First().ContractID;
            var Replacements = _db.AssetReplacements.Where(d => d.CustomerID == ClientID && d.ReplacementValue != null).ToList();
            if(Replacements != null)
            {
                 ReplacementValue = Replacements.Sum(c => c.ReplacementValue).Value;
            }
             else
            {
                  ReplacementValue = Replacements.Sum(c => c.ReplacementValue).Value;
            }
            
            var CostAvoidance =  String.IsNullOrEmpty(db.CostAvoidances.Where(c => c.CustomerID == ClientID).Sum(c => c.TotalSavingsCost).ToString()) ? "0" : db.CostAvoidances.Where(c => c.CustomerID == ClientID).Sum(c => c.TotalSavingsCost).Value.ToString();
            var Rollovers =   _db.QuarterlyRollovers.Where(r => r.ContractID == contract && r.Rollovers > 0).Sum(x => x.Rollovers * x.CPP) ?? 0;
           
            var Revision = er.RevisionSummary(contract).Sum(o=> o.Savings).ToString() ?? "0" ;

        
            var ret = new object[]
           {
                   new { label="Replacements", color = "#4acab4", data =   ReplacementValue },
                   new { label="REVision", color = "#ffea88" , data =  Revision},
                   new { label="Cost Avoidance", color = "#ff8153" , data =  CostAvoidance},
                   new { label="Rollovers", color = "#878bb6" , data =  Rollovers}
            };
            return Json(ret);
        }

        [HttpGet, Route("api/vitals/{id}")]
        public IHttpActionResult GetVitals(int id)
        {
            ExcelRevisionExport er = new ExcelRevisionExport();
            var contract = _context.vw_csContractList.OrderBy(c => c.ContractID).Where(c => c.CustomerID == id).First().ContractID;
            var Replacements = _db.AssetReplacements.Where(d => d.CustomerID == id).Sum(c => c.ReplacementValue) ?? 0;
            var CostAvoidance = _db.CostAvoidances.Where(c => c.CustomerID == id).Sum(c => c.TotalSavingsCost) ?? 0;
            var Rollovers =   _db.QuarterlyRollovers.Where(r => r.ContractID == contract && r.Rollovers > 0).Sum(x=> x.Rollovers * x.CPP);
         
            var Revision = er.RevisionSummary(contract).Sum(o => o.Savings);

            var ContractID = _context.SCContracts.FirstOrDefault(x => x.CustomerID == id).ContractID;
            var deviceCount =  _context.vw_admin_SCEquipments_22.Count(x => x.CustomerID == id && x.Active == true);
            var recovered = Replacements + CostAvoidance + Rollovers + Revision;
            var contractPages = ContractedPages(id);
            var loggedinUsers = _db.GlobalViewUsers.Count(x => x.idClient == id);
            return Json(new { devices = deviceCount,visitors = loggedinUsers, recovered, pages = contractPages });
        }
        [HttpGet, Route("api/getuserlogins/{id}")]
        public IHttpActionResult GetUserLogins(int id)
        {
            var userlist = _db.GlobalViewUsers.OrderByDescending(x=> x.logindatetime).Where(cu => cu.idClient == id)
                            .Select(cu => new {
                                Name = cu.FirstName + " " + cu.LastName,
                                Image = cu.userimage,
                                isLoggedIn = cu.isLoggedIn,
                                LoggedIn = cu.logindatetime,
                                LoggedOut = cu.logoutdatetime,
                                Duration = DbFunctions.DiffMinutes(cu.logindatetime, cu.logoutdatetime)
                            }).ToList();  

            return Json(userlist);
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

        public decimal? ContractedPages(int customerid)
        {

            var Volumes = _context.vw_ContractedPagesByDeviceTYpe_Cust.Where(x => x.CustomerID == customerid);
            return Volumes.Sum(x => x.Pages);
        }
        private decimal? visionSummary(int ContractID)
        {
            ExcelRevisionExport  re = new ExcelRevisionExport();
            var results = re.RevisionSummary(ContractID);
 
               var savings = results.Sum(s => s.Savings);
               var credits = results.Sum(s => s.Credits);
               return savings + credits;
            
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
