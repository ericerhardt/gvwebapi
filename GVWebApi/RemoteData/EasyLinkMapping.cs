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
    
    public partial class EasyLinkMapping
    {
        public int MappingID { get; set; }
        public Nullable<int> ChildId { get; set; }
        public Nullable<int> ClientId { get; set; }
        public string ClientName { get; set; }
        public Nullable<bool> IsEasyLinkOnly { get; set; }
    
        public virtual EasyLinkMapping EasyLinkMapping1 { get; set; }
        public virtual EasyLinkMapping EasyLinkMapping2 { get; set; }
    }
}
