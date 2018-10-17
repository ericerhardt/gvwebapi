using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GVWebapi.RemoteData;
namespace GVWebapi.Controllers
{
    public class ReportingController : ApiController
    {
        private readonly CoFreedomEntities _coFreedomEntities = new CoFreedomEntities();

        [HttpGet, Route("api/quarterlyreview/{ContractID}")]
        public IHttpActionResult QuarterlyReview(int ContractID)
        {
            var ContractStart = (from r in _coFreedomEntities.vw_csSCBillingContracts
                        where r.ContractID == ContractID && r.VoidFlag == 0
                        orderby r.InvoiceID descending
                        select new { value = r.OverageFromDate, text =  r.OverageFromDate }).ToList();

            var PeriodList = (from r in _coFreedomEntities.vw_csSCBillingContracts
                         where r.ContractID == ContractID && r.VoidFlag == 0
                         orderby r.InvoiceID ascending
                          select new { value = r.OverageToDate, text = r.OverageToDate }).ToList();
            return Json( new { StartList = ContractStart, PeriodList });

        }

    }
}
