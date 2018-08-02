using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class ServiceCallsViewModel
    {
        public string color { get; set; }
        public string label { get; set; }
        public IEnumerable<WeeklyCallTotals> data {get;set;}

    }
    public class WeeklyCallTotals
    {
        public string day { get; set; }
        public int totalcalls { get; set; }
    }
}