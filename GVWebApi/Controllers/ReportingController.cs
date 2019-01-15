using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;
using GVWebapi.Helpers;
using GVWebapi.Helpers.Reporting;
using GVWebapi.RemoteData;
using GVWebapi.Models.Reports;
using System.IO;
using System.Net.Http.Headers;

namespace GVWebapi.Controllers
{
    public class ReportingController : ApiController
    {
        private readonly CoFreedomEntities _coFreedomEntities = new CoFreedomEntities();
       

        [HttpGet, Route("api/volumetrendreport/{CustomerID}/{InvoiceID}")]
        public IHttpActionResult VolumeTrendReport(int CustomerID,int InvoiceID)
        { 

            var PeriodDates = (from r in _coFreedomEntities.vw_csSCBillingContracts
                              where r.InvoiceID == InvoiceID
                              select new {fromDate = r.OverageFromDate, toDate = r.OverageToDate, Contract = r.ContractID }).FirstOrDefault();
            var CustomerNumber = (from c in _coFreedomEntities.ARCustomers
                                  where c.CustomerID == CustomerID
                                  select c.CustomerNumber 
                                  ).FirstOrDefault();
          
            ExcelRevisionExport er = new ExcelRevisionExport();
            var results = er.GetVolumeTrend(CustomerNumber, PeriodDates.fromDate, PeriodDates.toDate).OrderBy(o => o.LineID);
            var groups = results.Select(x => x.MeterGroup).ToList().Distinct();
            return Json(new { volumetrend = results, metergroups = groups });

        }
        [HttpGet, Route("api/volumetrendperoids/{ContractID}")]
        public IHttpActionResult VolumeTrendPeroids(int ContractID)
        {
          

            var PeriodList = (from r in _coFreedomEntities.vw_csSCBillingContracts
                              where r.ContractID == ContractID && r.VoidFlag == 0
                              orderby r.InvoiceID descending
                              select new { value = r.InvoiceID, text = r.OverageToDate }).ToList();
            return Json(PeriodList);

        }


        [HttpGet, Route("api/quarterlyreview/{ContractID}")]
        public IHttpActionResult QuarterlyReview(int ContractID)
        {
            var ContractStart = (from r in _coFreedomEntities.vw_csSCBillingContracts
                        where r.ContractID == ContractID && r.VoidFlag == 0
                        orderby r.InvoiceID ascending
                        select new { value = r.OverageFromDate, text =  r.OverageFromDate }).ToList();

            var PeriodList = (from r in _coFreedomEntities.vw_csSCBillingContracts
                         where r.ContractID == ContractID && r.VoidFlag == 0
                         orderby r.InvoiceID descending
                          select new { value = r.OverageToDate, text = r.OverageToDate }).ToList();
            return Json( new { StartList = ContractStart, PeriodList });

        }
        [HttpPost, Route("api/downloadreport")]
        public IHttpActionResult DownloadReport(QuarterlyReviewModel model)
        {

            if (model != null)
            {
                var client = _coFreedomEntities.vw_CustomersOnContract.Where(c => c.CustomerID == model.CustomerID).FirstOrDefault();

                ExcelReport rep = new ExcelReport();

                string OverideDate = model.OverrideDate == null ? model.StartDate : model.OverrideDate;
                DateTime StartDate = DateTime.Parse(model.PeriodDate).AddMonths(-2);
                DateTime QTR = Convert.ToDateTime(model.PeriodDate);

                String filename = client.CustomerName + "-" + QTR.Month + "-" + QTR.Year + "_Quarterly.xlsx";
                string outfile = System.Web.Hosting.HostingEnvironment.MapPath("~/Reports/" + filename) ;

                //System.IO.FileInfo outputFile = new System.IO.FileInfo(outfile);
                //if (outputFile.Exists)
                //{
                //    outputFile.Delete();
                //    rep.CreatePackage(outfile, contractid, customer, perioddate);
                //}
                //else
                rep.CreatePackage(outfile, model.CustomerID, model.Customer, model.PeriodDate, model.ContractID,StartDate.ToShortDateString(), model.InvoiceID, OverideDate);
                //DownloadFile(outfile, filename);
                System.IO.FileInfo targetFile = new System.IO.FileInfo(outfile);


                if (!targetFile.Exists)
                {

                    return  BadRequest();
                }
                else
                {

                    try
                    {

                        //HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                        ////response.Content = new StreamContent(new FileStream(localFilePath, FileMode.Open, FileAccess.Read));
                        //Byte[] bytes = File.ReadAllBytes(outfile);
                        ////String file = Convert.ToBase64String(bytes);
                        //response.Content = new ByteArrayContent(bytes);
                        //response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                        //response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                        //response.Content.Headers.ContentDisposition.FileName = filename;

                        return Json(new { filename });

                    }
                    catch (Exception ex)
                    {

                        return InternalServerError();
                    }
                }
               
            }
            return  BadRequest();
        }
        protected HttpResponseMessage DownloadFile(string sFile, string filename)
        {

            System.IO.FileInfo targetFile = new System.IO.FileInfo(sFile);

       
                if (!targetFile.Exists)
                {
                
                   return Request.CreateResponse(HttpStatusCode.BadRequest);
                }
                else
                {

                    try
                    {

                    HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK);
                    //response.Content = new StreamContent(new FileStream(localFilePath, FileMode.Open, FileAccess.Read));
                    Byte[] bytes = File.ReadAllBytes(sFile);
                    //String file = Convert.ToBase64String(bytes);
                    response.Content = new ByteArrayContent(bytes);
                    response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
                    response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                    response.Content.Headers.ContentDisposition.FileName = filename;

                    return response;
 
                }
                catch (Exception ex)
                {

                    return Request.CreateResponse(HttpStatusCode.InternalServerError);
                }
            }
        }
       
    }
}
