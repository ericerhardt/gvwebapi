//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace GVWebapi.RemoteData
{
    using System;
    using System.Collections.Generic;
    
    public partial class Survey
    {
        public int SurveyID { get; set; }
        public int SurveyTypeID { get; set; }
        public int CustomerID { get; set; }
        public string CustomerName { get; set; }
        public string Name { get; set; }
        public string Title { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public Nullable<System.DateTime> SurveyDate { get; set; }
        public string SurveyComments { get; set; }
        public string Attachment { get; set; }
    }
}
