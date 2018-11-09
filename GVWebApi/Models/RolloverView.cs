using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public partial class RolloverView
    {
        public Nullable<int> ERPContractID { get; set; }
        public string ERPCustomerNumber { get; set; }
        public Nullable<System.DateTime> PeriodDate { get; set; }
        public Nullable<int> InvoiceID { get; set; }
        public string ERPMeterGroupDesc { get; set; }
        public Nullable<long> Rollover { get; set; }
        public Nullable<System.DateTime> StartDate { get; set; }
        public decimal CPP { get; set; }
        public Nullable<int> CustomerID { get; set; }
    }
}