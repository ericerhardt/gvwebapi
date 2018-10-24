using System;
using System.Xml; 
using System.Data.Entity;
using GVWebapi.Models;
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

        public List<GVWebapi.Models.VisionData> bindRevionDetailSummary(Int32 ContractID)
        {

            CoFreedomEntities  db = new CoFreedomEntities();
            var Company = (from c in db.vw_CustomersOnContract
                           join o in db.SCContracts on c.CustomerID equals o.CustomerID
                           where o.ContractID == ContractID
                           select c).First();
                     
             
  
           RevisionDataEntities _dbAudit = new RevisionDataEntities();
            var VisionDataList = new List<GVWebapi.Models.VisionData>();

            var revisions = (from r in _dbAudit.RevisionDataViews
                             where r.ERPContractID == ContractID
                             orderby r.PeriodDate descending
                select new
                {
                    ERPContarctID = ContractID,
                    PeriodDates =  r.PeriodDate.Value ,
                    StartDate = r.StartDate.Value,
                    PreFPRBaseExpense = r.ClientContractBase,
                    PreFPROverageExpense = r.OverageNoFPR,
                    FPRBaseExpense = r.FBRContractBase,
                    FPROverageExpense = r.OverageExpense,
                    Credits = r.Credits,
                    PeriodSavings = (r.ClientContractBase + r.OverageNoFPR) - (r.FBRContractBase + r.OverageExpense),
                    PctSavings =  r.NetSavings.Value != 0 && r.OverageNoFPR.Value!=0 ? r.NetSavings / (r.ClientContractBase + r.OverageNoFPR) : 0
                }).ToList();
                 var summary = (from s in revisions
                                group s by  new
                                {

                                    clientPeriodDates = s.PeriodDates ,
                                    clientStartDate = s.StartDate,
                                    fprCost =s.FPRBaseExpense,
                                    clientCost = s.PreFPRBaseExpense 
                                }
                                    into v
                                    select new
                                    {
              
                                        clientPeriodDates  = v.Key.clientPeriodDates,
                                        clientStartDate    = v.Key.clientStartDate,
                                        fprOverageCost     = v.Sum(o => o.FPROverageExpense),
                                        fprCost            = v.Key.fprCost,
                                        clientOverageCost  = v.Sum(o => o.PreFPROverageExpense),
                                        clientCost         = v.Key.clientCost,
                                        credits            = v.Sum(o => o.Credits),
                                        savings            = v.Sum(o => o.PeriodSavings) ,
                                        pct                = v.Sum(o => o.PctSavings)
                                    }).OrderByDescending(o => o.clientPeriodDates);


                 foreach (var revision in summary)
                 {
                     Int32 monthsDiff = TotelMonthDifference(revision.clientPeriodDates, revision.clientStartDate);
                     var a = new GVWebapi.Models.VisionData();
                     a.ERPContractID = ContractID;
                     a.ClientStartDate = revision.clientStartDate;
                     a.ClientPeriodDates = revision.clientPeriodDates.AddMonths(-monthsDiff).ToString("MMM") + " - " + revision.clientPeriodDates.ToString("MMM yyyy");
                     a.FPROverageCost = Convert.ToDouble(revision.fprOverageCost);
                     a.FPRCost = Convert.ToDouble(revision.fprCost) * (monthsDiff +1);
                     a.ClientOverageCost = Convert.ToDouble(revision.clientOverageCost);
                     a.ClientCost = Convert.ToDouble(revision.clientCost) * (monthsDiff + 1);
                     a.Credits = Convert.ToDouble(revision.credits);
                     a.Savings = Convert.ToDouble(revision.savings);
                     a.Pct = Convert.ToDouble(revision.pct);
                     VisionDataList.Add(a);
                 }
         return VisionDataList;
    }

        public List<VisionDataDetail> bindRevisionLegecy(Int32 ContractID)
        {
            using (var _dbAudit = new RevisionDataEntities())
            {
                var meterGroups = (from mg in _dbAudit.MeterGroups
                                   where mg.ERPContractID == ContractID
                                   select mg).ToList();

                var clientdata = (from cd in _dbAudit.ClientDatas
                                  where cd.ERPContractID == ContractID
                                  select cd).First();
                var billID = clientdata.BilledID;
                Decimal FBRBase = 0;
                Decimal ClientBase = 0;
                if (billID == 0)
                {

                    FBRBase = Convert.ToDecimal(clientdata.BaseContractAmount * 3);
                    ClientBase = Convert.ToDecimal(clientdata.AffectedCost * 3);
                }
                else
                {
                    FBRBase = Convert.ToDecimal(clientdata.BaseContractAmount);
                    ClientBase = Convert.ToDecimal(clientdata.AffectedCost);
                }
                var VisionDataList = new List<VisionDataDetail>();

                RevisionDataEntities ad = new RevisionDataEntities();
                var legecyRecords = (from r in ad.RevisionDataViews
                                     where r.ERPContractID == ContractID orderby r.PeriodDate descending
                                     select new
                                     {
                                         ContractID = r.ERPContractID,
                                         MeterGroup = r.ERPMeterGroupDesc,
                                         ContractVolume = r.ContractedVolume.HasValue ? r.ContractedVolume.Value : 0,
                                         ActualVolume = r.ActualVolume.HasValue ? r.ActualVolume.Value : 0,
                                       //  VolumeOffset = r.VolumeOffset != null  ? r.VolumeOffset : 0,
                                         CPP = r.CPPRate,
                                         ClientCPP = r.ClientCPP,
                                         Rollover = r.Rollover,
                                         CreditAmount = r.Credits.HasValue ? r.Credits.Value :0,
                                  //       AdjustedVolume = r.AdustedContractVolume.HasValue ? r.AdustedContractVolume.Value : 0,
                                         OverageToDate = r.PeriodDate,
                                         OverageFromDate = r.StartDate,
                                         OverageCharge = r.OverageExpense.HasValue ? r.OverageExpense.Value : 0,
                                         clientOverage = r.OverageNoFPR.HasValue ? r.OverageNoFPR.Value : 0,
                                         clientSavings = r.NetSavings.HasValue ? r.NetSavings.Value : 0,
                                         PctSavings = 0//r.NetSavings.HasValue && r.NetSavings != 0 ? r.NetSavings.Value / r.OverageNoFPR.Value : 0

                                     }).Distinct().OrderByDescending(o => o.OverageToDate);
                foreach (var legecy in legecyRecords)
                {
                    VisionDataDetail visiondata = new VisionDataDetail();
                    visiondata.ERPContractID = legecy.ContractID.Value;
                    visiondata.MeterGroup = legecy.MeterGroup;
                    visiondata.ClientPeriodDates = PeriodFormatDate(legecy.OverageToDate.Value,legecy.OverageFromDate.Value);
                    visiondata.PeriodDate = legecy.OverageToDate.Value;
                  //  visiondata.AdjustedVolume = legecy.AdjustedVolume;
                    visiondata.ActualVolume = legecy.ActualVolume;
                  //  visiondata.VolumeOffset = legecy.VolumeOffset;
                    visiondata.ContractVolume = legecy.ContractVolume;
                    visiondata.clientOverage = Convert.ToDouble(legecy.clientOverage);
                    visiondata.OverageCost = Convert.ToDouble(legecy.OverageCharge) - Convert.ToDouble(legecy.CreditAmount);
                    visiondata.OverageVolume = Convert.ToInt32(legecy.ActualVolume - legecy.ContractVolume);
                    visiondata.month = legecy.OverageToDate.Value.ToString("{0:MMMM}");
                    visiondata.RolloverVolume = legecy.Rollover.Value;
                    visiondata.CPP = Convert.ToDecimal(legecy.CPP.Value);
                    visiondata.ClientCPP = Convert.ToDouble(legecy.ClientCPP.Value);
                    visiondata.CreditAmount = Convert.ToDouble(legecy.CreditAmount);
                    visiondata.Savings = Convert.ToDouble(legecy.clientSavings);
                    visiondata.PctSavings = Convert.ToDouble(legecy.PctSavings);
                    VisionDataList.Add(visiondata);



                }
                return VisionDataList;
            }     
          
        }
        public String GetPeroidNotes(DateTime PeriodDates, int ContractID)
        {
             RevisionDataEntities ad = new RevisionDataEntities();
             var query = (from p in ad.PeriodNotes
                          where p.ContractID == ContractID && p.PeriodDate == PeriodDates
                          select p.PeriodNote1).FirstOrDefault();
             return query;
        }
    
      
    }
}
