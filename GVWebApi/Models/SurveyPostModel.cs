using GVWebapi.RemoteData;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace GVWebapi.Models
{
    public class SurveyPostModel
    {
        public int SurveyID { get; set; }
      
        public int CustomerID { get; set; }
        public string CustomerName { get; set; }
        public string Name { get; set; }
        public string Title { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public Nullable<System.DateTime> SurveyDate { get; set; }
       
        public IList<SurveyAnswer> Answers { get; set; }
    }
}