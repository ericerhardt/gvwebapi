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
            Int32 monthsDiff = TotalMonthDifference(_date, _startdate);
            String StartDate = String.Format("{0:MMMM yyyy}", _date.AddMonths(-monthsDiff));
            String EndDate = String.Format("{0:MMMM yyyy}", _date);
            return StartDate + " - " + EndDate;
          
        }
        private Int32 TotalMonthDifference(DateTime dtThis, DateTime dtOther)
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
        
        public List<GVWebapi.Models.VisionData>  RevisionSummary(Int32 ContractID)
        {

            GlobalViewEntities gv = new GlobalViewEntities();
            CoFreedomEntities ea = new CoFreedomEntities();
            List<RevisionHistoryModel> query = new List<RevisionHistoryModel>();
          
            var VisionDataList = new List<GVWebapi.Models.VisionData>();
            List<RevisionHistoryModel> RevisionModel = new List<RevisionHistoryModel>();
            
            var gvrevision = gv.RevisionDataViews.Where(r => r.ContractID == ContractID).ToList();
            var gvexpense = gv.RevisionBaseExpenses.Where(r => r.ContractID == ContractID).OrderByDescending(r=> r.OverrideDate).ToList();
            for(var i =0; i < gvexpense.Count; i++)  
            {
                if (i == 0)
                    gvexpense[i].EndDate = DateTime.Now;
                else
                gvexpense[i].EndDate = gvexpense[i - 1].OverrideDate.Value.AddDays(-1);
            }

            var gvmetergroups = gv.RevisionMeterGroups.Where(r => r.ERPContractID == ContractID && r.rollovers == true).ToList();
            var revisions = from e in gvrevision
                          
                            orderby e.MeterGroup
                            select
                            new
                            {
                               // FPROverageCharge =  CalcuateOverageCharge( e.ActualVolume.Value ,e.ContractVolume.Value , e.CPP.Value, e.Rollover),
                                //ClientOverageCharge = CalcuateOverageCharge(e.ActualVolume.Value, e.ContractVolume.Value, mg.CPP, e.Rollover),
                                FPROverageCharge = ((e.Overage - e.Rollover) * e.CPP) - e.CreditAmount <= 0 ? 0.00M : (((e.Overage) - e.Rollover) * e.CPP) - e.CreditAmount,
                                //ClientOverage = (((e.ActualVolume - e.Rollover) - e.ContractVolume) * mg.CPP) - e.CreditAmount <= 0 ? 0.00M : (((e.ActualVolume - e.Rollover) - e.ContractVolume) * mg.CPP) - e.CreditAmount,
                                ClientOverageCharge = (( e.Overage) * e.CPP) - e.CreditAmount <= 0 ? 0.00M : (((e.Overage)) * e.CPP) - e.CreditAmount,
                                CreditAmount = e.CreditAmount,
                                OverageToDate = e.OverageToDate,
                                OverageFromDate = e.OverageFromDate,
                                ContractID = e.ContractID,
                                ContractMeterGroupID = e.ContractMeterGroupID
                            };

           var startDate = gvrevision.Min(o=> o.OverageFromDate);
            var endDate = gvrevision.Max(o => o.OverageToDate);

            var months = MonthsBetween(startDate.Value, endDate.Value);

            var monthlyCost = new List<Tuple<decimal,decimal,DateTime>>();
            foreach(var month in months)
            {
                var oldcost = gvexpense.Where(o => month.Item1 >= o.OverrideDate.Value && month.Item1 <= o.EndDate).Select(o => o.PreBase).FirstOrDefault();
                var newcost = gvexpense.Where(o => month.Item1 >= o.OverrideDate.Value && month.Item1 <= o.EndDate).Select(o => o.FprBase).FirstOrDefault();
                if(oldcost != null && newcost != null)
                {

                
                var Expenses = Tuple.Create(oldcost.Value, newcost.Value,month.Item1);
                monthlyCost.Add(Expenses);
                }
            }


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
                                fprOverageCost = v.Sum(o => o.FPROverageCharge),
                               fprBase = monthlyCost.Where( o=> o.Item3 >= v.Key.clientStartDate.Value && o.Item3 <= v.Key.clientPeriodDates.Value).Sum(o => o.Item2),
                               clientBase = monthlyCost.Where(o => o.Item3 >= v.Key.clientStartDate.Value && o.Item3 <= v.Key.clientPeriodDates.Value).Sum(o => o.Item1),
                                clientOverageCost = v.Sum(o => o.ClientOverageCharge),
                               credits = v.Sum(o=> o.CreditAmount),
                           }).OrderByDescending(o => o.clientPeriodDate).ToList();


            foreach (var revision in summary)
            {
                Int32 monthsDiff = TotalMonthDifference(revision.clientPeriodDate, revision.clientStartDate);
               var fprBase = gvexpense.Where(o => o.ContractID == revision.contractId &&  (revision.clientPeriodDate >= o.OverrideDate && revision.clientStartDate <= o.EndDate)).Select(o => o.FprBase).ToList();
               var clientBase = gvexpense.Where(o => o.ContractID == revision.contractId && (revision.clientPeriodDate >= o.OverrideDate && revision.clientStartDate <= o.EndDate)).Select(o => o.PreBase).ToList();
                var a = new GVWebapi.Models.VisionData();
                a.ERPContractID = ContractID;
                a.ClientStartDate = revision.clientStartDate;
                a.ClientPeriodDate = revision.clientPeriodDate;
                a.ClientPeriodDates = revision.clientPeriodDate.AddMonths(-monthsDiff).ToString("MMM") + " - " + revision.clientPeriodDate.ToString("MMM yyyy");
                 a.FPROverageCost =  revision.fprOverageCost.Value;
               
                a.FPRCost =  revision.fprBase;
                a.ClientOverageCost =  revision.clientOverageCost.Value;
                 
                a.ClientCost = revision.clientBase;
                a.Credits =  revision.credits * (monthsDiff + 1);
                a.Savings =  (a.ClientCost + a.ClientOverageCost) - ((a.FPRCost + a.FPROverageCost ) - a.Credits);
                if(a.Savings != 0.00M)
                a.Pct =  (a.Savings /(a.ClientCost + a.ClientOverageCost));
                if (a.Pct < 0)
                    a.Pct = 0.00M;
                VisionDataList.Add(a);
            }
            return VisionDataList;
        }
         
        public decimal CalcuateOverageCharge(decimal ActualVolume, decimal ContractVolume, decimal CPP, int? Rollovers)
        {
            var _rollovers = Rollovers == null ? 0 : Rollovers;
            var overage = (ContractVolume + _rollovers) - ActualVolume;
            if(overage > 0)
            {
                return overage.Value * CPP;
            }
            return 0.00M;
        }

        public IEnumerable<RevisionHistoryModel> GetRevisionHistory(int ContractID)
        {
            GlobalViewEntities gv = new GlobalViewEntities();
            CoFreedomEntities ea = new CoFreedomEntities();
            List<RevisionHistoryModel> query = new List<RevisionHistoryModel>();
            IEnumerable<vw_REVisionInvoices> periods = ea.vw_REVisionInvoices.OrderByDescending(p => p.InvoiceID).Where(p => p.ContractID == ContractID).ToList();
            List<RevisionHistoryModel> RevisionModel = new List<RevisionHistoryModel>();
             
            var gvdata = gv.RevisionDataViews.Where(r => r.ContractID == ContractID).ToList().Distinct();

            var query2 = from   g in gvdata
                         orderby g.MeterGroup
                         select  
 
                new RevisionDataModel
            {
                RevisionID = g.RevisionID,  
                InvoiceID = g.InvoiceID,
                ContractMeterGroupID = g.ContractMeterGroupID,
                MeterGroup = g.MeterGroup,
                ContractVolume = g.ContractVolume,
                ActualVolume = g.ActualVolume,
                Overage =  g.ActualVolume - (g.ContractVolume + g.Rollover) ,
                CPP = g.CPP,
                OverageCharge = ((g.ActualVolume - (g.ContractVolume + g.Rollover)) * g.CPP ) - g.CreditAmount <= 0 ? 0.00M : ((g.ActualVolume -(g.ContractVolume + g.Rollover)) * g.CPP) - g.CreditAmount,
                CreditAmount = g.CreditAmount,
                Rollover = g.Rollover,
                OverageToDate = g.OverageToDate,
                OverageFromDate = g.OverageFromDate,
                ContractID = g.ContractID,

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
        public IEnumerable<VisionData> GetRevisionSummary(int ContractID)
        {
            GlobalViewEntities gv = new GlobalViewEntities();
            CoFreedomEntities ea = new CoFreedomEntities();
           
            IEnumerable<vw_REVisionInvoices> periods = ea.vw_REVisionInvoices.OrderByDescending(p => p.InvoiceID).Where(p => p.ContractID == ContractID).ToList();
            List<RevisionHistoryModel> RevisionModel = new List<RevisionHistoryModel>();
            var VisionDataList = new List<VisionData>();
            var gvdata = gv.RevisionDataViews.Where(r => r.ContractID == ContractID).ToList().Distinct();

            var gvexpense = gv.RevisionBaseExpenses.Where(r => r.ContractID == ContractID).OrderByDescending(r => r.OverrideDate).ToList();
            for (var i = 0; i < gvexpense.Count; i++)
            {
                if (i == 0)
                    gvexpense[i].EndDate = DateTime.Now;
                else
                    gvexpense[i].EndDate = gvexpense[i - 1].OverrideDate.Value.AddDays(-1);
            }

            var gvmetergroups = gv.RevisionMeterGroups.Where(r => r.ERPContractID == ContractID && r.rollovers == true).ToList();

            var query2 = from g in gvdata
                         join mg in gvmetergroups
                           on g.ContractMeterGroupID equals mg.ERPMeterGroupID
                         where mg.rollovers == true
                         orderby g.MeterGroup
                         select

                new RevisionDataModel
                {

                    InvoiceID = g.InvoiceID,
                    ContractMeterGroupID = g.ContractMeterGroupID,
                    MeterGroup = g.MeterGroup,
                    ContractVolume = g.ContractVolume,
                    ActualVolume = g.ActualVolume,
                    Overage = g.ActualVolume - (g.ContractVolume + g.Rollover),
                    CPP = g.CPP,
                    ClientCPP = mg.CPP,
                    OverageCharge = ((g.ActualVolume - (g.ContractVolume + g.Rollover)) * g.CPP) - g.CreditAmount <= 0 ? 0.00M : ((g.ActualVolume - (g.ContractVolume + g.Rollover)) * g.CPP) - g.CreditAmount,
                    CreditAmount = g.CreditAmount,
                    Rollover = g.Rollover,
                    OverageToDate = g.OverageToDate,
                    OverageFromDate = g.OverageFromDate,
                    ContractID = g.ContractID,

                };
              
            foreach (var period in periods)
            {
                VisionData model = new VisionData();
                var results = query2.Where(x => x.InvoiceID == period.InvoiceID);
                Int32 monthsDiff = TotalMonthDifference(period.PeriodDate.Value, period.StartDate.Value);
                var fprBase = gvexpense.Where(o => o.ContractID == period.ContractID &&  (period.PeriodDate >= o.OverrideDate && period.StartDate <= o.EndDate)).Select(o => o.FprBase).ToList();
                var clientBase = gvexpense.Where(o => o.ContractID == period.ContractID && (period.PeriodDate >= o.OverrideDate && period.StartDate <= o.EndDate)).Select(o => o.PreBase).ToList();

                model.ERPContractID = ContractID;
                model.ClientStartDate = period.StartDate.Value;
                model.ClientPeriodDate = period.PeriodDate.Value;
                model.ClientPeriodDates = period.StartDate.Value.ToString("MMM") + " - " + period.PeriodDate.Value.ToString("MMM yyyy");
                model.FPROverageCost = results.Sum(x=> x.OverageCharge.Value);
                model.FPRCost =  fprBase.FirstOrDefault().Value;
                model.ClientOverageCost = results.Sum(x => x.Overage.Value * x.ClientCPP.Value);
                model.ClientCost = clientBase.FirstOrDefault().Value;
                model.Credits = results.Sum(x => x.CreditAmount.Value);
                model.Savings = (model.ClientCost + model.ClientOverageCost) - ((model.FPRCost + model.FPROverageCost) - model.Credits);
                if (model.Savings != 0.00M)
                    model.Pct = (model.Savings / (model.ClientCost + model.ClientOverageCost));
                if (model.Pct < 0)
                    model.Pct = 0.00M;

                VisionDataList.Add(model);
            }

            return VisionDataList;
        }
        public IEnumerable<RolloverPagesModel> GetRolloverHistory(int ContractID)
        {
            GlobalViewEntities gv = new GlobalViewEntities();
            CoFreedomEntities ea = new CoFreedomEntities();
            List<RolloverPagesModel> rollOverModels = new List<RolloverPagesModel>();
        
            
         
            IEnumerable<vw_REVisionInvoices> periods = ea.vw_REVisionInvoices.OrderByDescending(p => p.InvoiceID).Where(p => p.ContractID == ContractID).ToList();
            foreach (var period in periods)
            {
                RolloverPagesModel rollOverPage = new RolloverPagesModel();
                rollOverPage.Period = period.PeriodDate;
                rollOverPage.InvoiceID = period.InvoiceID;
                rollOverPage.StartDate = period.StartDate;
                var gvdata = gv.QuarterlyRollovers.Where(r => r.ContractID == ContractID && r.InvoiceID == period.InvoiceID).ToList();
                var _rollovers = from g in gvdata
                                 orderby g.ERPMeterGroupDesc
                                 select
                                 new Rollovers
                                 {
                                    InvoiceID = g.InvoiceID,
                                    StartDate = period.StartDate,
                                    ERPMeterGroupDesc = g.ERPMeterGroupDesc,
                                    CPP = g.CPP,
                                    Rollover = g.Rollovers,
                                    Savings =  g.Rollovers * g.CPP
                                 };
                  

               
               
                rollOverPage.TotalSavings = _rollovers.Sum(r => r.Savings).Value;
                rollOverPage.Data = _rollovers.Distinct();
                rollOverModels.Add(rollOverPage);
            }

            return rollOverModels ;
        }
        public IEnumerable<VolumeTrendModel> GetVolumeTrend(string CustomerNumber,string StartDate,string PeriodDate)
        {
            CoFreedomEntities db = new CoFreedomEntities();
            var ToDate = Convert.ToDateTime(PeriodDate).AddDays(1);
            var FromDate = Convert.ToDateTime(StartDate).AddDays(1);
            var vtrends = db.Database.SqlQuery<VolumeTrendModel>("exec csVolumeTrend @vd_FromDate, @vd_ToDate, @vs_Customer, @vs_CustomerNumber", new SqlParameter("@vd_FromDate", FromDate), new SqlParameter("@vd_ToDate", ToDate), new SqlParameter("@vs_Customer", ""), new SqlParameter("@vs_CustomerNumber", CustomerNumber)).ToList();
            foreach(var trend in vtrends)
            {
                if (trend.LastPeriodVolume > 0 && trend.PeriodVolume > 0)
                {
                    trend.VolumeDiff = ((trend.LastPeriodVolume - trend.PeriodVolume) / trend.LastPeriodVolume) * -100;
                }
                else
                {
                    trend.VolumeDiff = 0.0M;
                }
            }
            return vtrends;
        }
        public static IEnumerable<Tuple<DateTime>> MonthsBetween(DateTime startDate, DateTime endDate)
        {
            DateTime iterator;
            DateTime limit;

            if (endDate > startDate)
            {
                iterator = new DateTime(startDate.Year, startDate.Month, 1);
                limit = endDate;
            }
            else
            {
                iterator = new DateTime(endDate.Year, endDate.Month, 1);
                limit = startDate;
            }

            var dateTimeFormat = CultureInfo.CurrentCulture.DateTimeFormat;
            while (iterator <= limit)
            {
                yield return Tuple.Create(iterator);
                iterator = iterator.AddMonths(1);
            }
        }
    }
}
