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
    
    public partial class EasyLinkItem
    {
        public long EasyLinkItemId { get; set; }
        public long EasyLinkId { get; set; }
        public int Child { get; set; }
        public string Email { get; set; }
        public string Fax { get; set; }
        public System.DateTime DateTime { get; set; }
        public string Description { get; set; }
        public decimal Duration { get; set; }
        public int Pages { get; set; }
        public decimal Charge { get; set; }
        public string MessageNumber { get; set; }
    
        public virtual EasyLink EasyLink { get; set; }
    }
}
