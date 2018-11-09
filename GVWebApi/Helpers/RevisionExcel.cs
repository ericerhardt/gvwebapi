using System;
using System.Xml; 
using System.Data.Entity;
using GVWebapi.Models;
using System.Data.SqlClient;
using GVWebapi.RemoteData;
using System.Globalization;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace GVWebapi.Helpers
{    
    public class ExcelRevisionExport {

        public bool _newContract;
        private Int32 _contractID;
        private String _customerName = String.Empty;

       
        private string PeriodFormatDate(DateTime _date,DateTime _startdate)
        {
            _customerName = "";
            Int32 monthsDiff = TotelMonthDifference(_date, _startdate);
            String StartDate = String.Format("{0:MMMM yyyy}", _date.AddMonths(-monthsDiff));
            String EndDate = String.Format("{0:MMMM yyyy}", _date);
            return StartDate + " - " + EndDate;
          
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
        public Int32 GetContractID(int CustomerID)
        {
            CoFreedomEntities ea = new CoFreedomEntities();
            var Contract = ea.vw_ClientsOnContract.Where(c => c.CustomerID == CustomerID).FirstOrDefault();
            return Contract.ContractID;
        }
        
        public List<GVWebapi.Models.VisionData> bindRevionDetailSummary(Int32 ContractID)
        {

            GlobalViewEntities gv = new GlobalViewEntities();
            CoFreedomEntities ea = new CoFreedomEntities();
            List<RevisionHistoryModel> query = new List<RevisionHistoryModel>();
          
            var VisionDataList = new List<GVWebapi.Models.VisionData>();
            List<RevisionHistoryModel> RevisionModel = new List<RevisionHistoryModel>();
            var eadata = ea.vw_RevisionInvoiceHistory.Where(r => r.ContractID == ContractID).ToList();
            var gvrevision = gv.RevisionDatas.Where(r => r.ContractID == ContractID).ToList();
            var gvexpense = gv.RevisionBaseExpenses.Where(r => r.ContractID == ContractID).ToList();
            var gvmetergroups = gv.RevisionMeterGroups.Where(r => r.ERPContractID == ContractID).ToList();
            var revisions = from e in eadata
                         join g in gvrevision
                         on new { e.InvoiceID, MeterGroup = e.ContractMeterGroupID.Value } equals new { g.InvoiceID, MeterGroup = g.MeterGroupID }
                         join mg in gvmetergroups
                         on new {eContractID = e.ContractID, MeterGroup = e.ContractMeterGroupID.Value } equals new { eContractID = mg.ERPContractID.Value, MeterGroup = mg.ERPMeterGroupID.Value }
                            orderby e.ContractMeterGroup
                         select

                new 
                {
                    FPROverage = (((e.ActualVolume - g.Rollover) - e.ContractVolume) * e.CPP) - g.Credits <= 0 ? 0.00M : (((e.ActualVolume - g.Rollover) - e.ContractVolume) * e.CPP) - g.Credits,                 
                    ClientOverage = (((e.ActualVolume - g.Rollover) - e.ContractVolume) * mg.CPP) - g.Credits <= 0 ? 0.00M : (((e.ActualVolume - g.Rollover) - e.ContractVolume) * mg.CPP) - g.Credits,
                    CreditAmount = g.Credits,
                    OverageToDate = e.OverageToDate,
                    OverageFromDate = e.OverageFromDate,
                    ContractID = e.ContractID,
                    ContractMeterGroupID = e.ContractMeterGroupID
                };

           
            
            var summary = (from r in revisions
                           group r by new
                           {
                               clientPeriodDates =r.OverageToDate ,
                               clientStartDate = r.OverageFromDate ,
                               contractId = r.ContractID
                           }
                               into v 
                           select new
                           {
                               contractId = v.Key.contractId,
                               clientPeriodDate = v.Key.clientPeriodDates.Value,
                               clientStartDate = v.Key.clientStartDate.Value,  
                               fprOverageCost = v.Sum(o => o.FPROverage),
                               
                               clientOverageCost = v.Sum(o => o.ClientOverage),
                               credits = v.Sum(o=> o.CreditAmount),
                           }).OrderByDescending(o => o.clientPeriodDate).ToList();


            foreach (var revision in summary)
            {
                Int32 monthsDiff = TotelMonthDifference(revision.clientPeriodDate, revision.clientStartDate);
                var fprBase = gvexpense.Where(o => o.ContractID == revision.contractId && o.OverrideDate <= revision.clientStartDate).Select(o => o.FprBase).FirstOrDefault();
                var clientBase = gvexpense.Where(o => o.ContractID == revision.contractId &&  o.OverrideDate <= revision.clientStartDate).Select(o => o.PreBase).FirstOrDefault();
                var a = new GVWebapi.Models.VisionData();
                a.ERPContractID = ContractID;
                a.ClientStartDate = revision.clientStartDate;
                a.ClientPeriodDates = revision.clientPeriodDate.AddMonths(-monthsDiff).ToString("MMM") + " - " + revision.clientPeriodDate.ToString("MMM yyyy");
                a.FPROverageCost =  revision.fprOverageCost.Value;
                a.FPRCost =  fprBase.Value  *  (monthsDiff + 1);
                a.ClientOverageCost =  revision.clientOverageCost.Value;
                a.ClientCost =  clientBase.Value  *  (monthsDiff + 1);
                a.Credits =  revision.credits.Value * (monthsDiff + 1);
                a.Savings =  (a.ClientCost + a.ClientOverageCost) - ((a.FPRCost + a.FPROverageCost ) - a.Credits);
                a.Pct =  (a.Savings /(a.ClientCost + a.ClientOverageCost));
                VisionDataList.Add(a);
            }
            return VisionDataList;
        }

      

        public IEnumerable<RevisionHistoryModel> GetRevisionHistory(int ContractID)
        {
            GlobalViewEntities gv = new GlobalViewEntities();
            CoFreedomEntities ea = new CoFreedomEntities();
            List<RevisionHistoryModel> query = new List<RevisionHistoryModel>();
            IEnumerable<vw_REVisionInvoices> periods = ea.vw_REVisionInvoices.OrderByDescending(p => p.InvoiceID).Where(p => p.ContractID == ContractID).ToList();
            List<RevisionHistoryModel> RevisionModel = new List<RevisionHistoryModel>();
            var eadata = ea.vw_RevisionInvoiceHistory.Where(r =>  r.ContractID == ContractID).ToList();
            var gvdata = gv.RevisionDatas.Where(r => r.ContractID == ContractID).ToList();

            var query2 = from e in eadata
                         join g in gvdata
                         on new { e.InvoiceID, MeterGroup = e.ContractMeterGroupID.Value } equals new { g.InvoiceID, MeterGroup  = g.MeterGroupID }
                         orderby e.MeterGroup
                         select  
 
                new RevisionDataModel
            {
                RevisionID = g.RevisionID,
                InvoiceID = e.InvoiceID,
                ContractMeterGroupID = e.ContractMeterGroupID,
                MeterGroup = e.MeterGroup,
                ContractVolume = e.ContractVolume,
                ActualVolume = e.ActualVolume,
                Overage = e.ActualVolume - (e.ContractVolume + g.Rollover) ,
                CPP = e.CPP,
                OverageCharge = ((e.ActualVolume - (e.ContractVolume + g.Rollover)) * e.CPP ) - g.Credits <= 0 ? 0.00M : ((e.ActualVolume -(e.ContractVolume + g.Rollover)) * e.CPP) - g.Credits,
                CreditAmount = g.Credits,
                Rollover = g.Rollover,
                OverageToDate = e.OverageToDate,
                OverageFromDate = e.OverageFromDate,
                ContractID = e.ContractID,

            };
            

            foreach (var period in periods)
            {
                RevisionHistoryModel model = new RevisionHistoryModel();

                model.peroid =  period.PeriodDate;
                model.InvoiceId = period.InvoiceID;
                model.detail = query2.Where(x => x.InvoiceID == period.InvoiceID);
                model.Notes = gv.RevisionNotes.Where(x => x.ContractID == period.ContractID && x.InvoiceID == period.InvoiceID).Select(x => x.Notes).FirstOrDefault();
                RevisionModel.Add(model);
            }

            return RevisionModel;
        }
        public IEnumerable<RolloverPagesModel> GetRolloverHistory(int ContractID)
        {
            GlobalViewEntities gv = new GlobalViewEntities();
            CoFreedomEntities ea = new CoFreedomEntities();
            List<RolloverPagesModel> rollOverModels = new List<RolloverPagesModel>();
            var eadata = gv.RevisionDatas.Where(r => r.ContractID == ContractID).ToList();
            var gvdata = gv.RevisionMeterGroups.Where(r => r.ERPContractID == ContractID).ToList();
         
            IEnumerable<vw_REVisionInvoices> periods = ea.vw_REVisionInvoices.OrderByDescending(p => p.InvoiceID).Where(p => p.ContractID == ContractID).ToList();
            foreach (var period in periods)
            {
                RolloverPagesModel rollOverPage = new RolloverPagesModel();
                rollOverPage.Period = period.PeriodDate;
                rollOverPage.InvoiceID = period.InvoiceID;
                rollOverPage.StartDate = period.StartDate;
                var _rollovers = from e in eadata
                                 join g in gvdata
                                 on new { MeterGroup = e.MeterGroupID } equals new {  MeterGroup = g.ERPMeterGroupID.Value }
                                 where e.InvoiceID == period.InvoiceID
                                 orderby e.MeterGroupID
                                 select
                                 new Rollovers
                                 {
                                    InvoiceID = e.InvoiceID,
                                    StartDate = period.StartDate,
                                    ERPMeterGroupDesc = g.ERPMeterGroupDesc,
                                    CPP = g.CPP,
                                    Rollover = e.Rollover,
                                    Savings = (e.Rollover * g.CPP)
                                 };
                  

               
               
                rollOverPage.TotalSavings = _rollovers.Sum(r => r.Savings).Value;
                rollOverPage.Data = _rollovers;
                rollOverModels.Add(rollOverPage);
            }

            return rollOverModels ;
        }
        public IEnumerable<VolumeTrendModel> GetVolumeTrend(string CustomerNumber,string StartDate,string PeriodDate)
        {
            CoFreedomEntities db = new CoFreedomEntities();
            var ToDate = Convert.ToDateTime(PeriodDate).AddDays(1);
            var FromDate = Convert.ToDateTime(StartDate).AddDays(1);
            var vtrend = db.Database.SqlQuery<VolumeTrendModel>("exec csVolumeTrend @vd_FromDate, @vd_ToDate, @vs_Customer, @vs_CustomerNumber", new SqlParameter("@vd_FromDate", FromDate), new SqlParameter("@vd_ToDate", ToDate), new SqlParameter("@vs_Customer", ""), new SqlParameter("@vs_CustomerNumber", CustomerNumber)).ToList();

            return vtrend;
        }
    }
}
