using System;
using System.Collections.Generic;
using GVWebapi.Services;

namespace GVWebapi.Models.Reconciliation
{
    public class ReconciliationViewModel
    {
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public decimal  ReconcileAdj { get; set; }
        public IList<CycleSummaryModel> CycleSummary { get; set; } = new List<CycleSummaryModel>();
        public IList<InvoicedServiceModel> InvoicedService { get; set; } = new List<InvoicedServiceModel>();
        public IList<CostByDeviceModel> CostByDevice { get; set; } = new List<CostByDeviceModel>();
    }
}