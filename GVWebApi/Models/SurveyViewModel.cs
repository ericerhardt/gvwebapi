using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using GVWebapi.RemoteData;
namespace GVWebapi.Models
{
    public class SurveyViewModel
    {
        public SurveyWithAvg Survey { get; set; }
        public IEnumerable<SurveyQuestionsWithAnswer> SurveyDetail { get; set; }
    
     
    }
}