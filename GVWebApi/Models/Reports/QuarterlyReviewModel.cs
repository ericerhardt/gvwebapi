using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models.Reports
{
    public class QuarterlyReviewModel
    {
        public int ContractID { get; set; }
        public int CustomerID { get; set; }
        public int InvoiceID { get; set; }
        public string Customer { get; set; }
        public string PeriodDate { get; set; }
        public string StartDate { get; set; }
        public string OverrideDate { get; set; }
    }
}
 