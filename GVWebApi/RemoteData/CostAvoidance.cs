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
    
    public partial class CostAvoidance
    {
        public int CostAvoidanceID { get; set; }
        public int CustomerID { get; set; }
        public string Location { get; set; }
        public Nullable<System.DateTime> SavingsDate { get; set; }
        public Nullable<System.DateTime> EndDate { get; set; }
        public Nullable<int> SavingsType { get; set; }
        public Nullable<decimal> SavingsCost { get; set; }
        public Nullable<decimal> TotalSavingsCost { get; set; }
        public Nullable<int> Months { get; set; }
        public string Comments { get; set; }
    }
}
