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
    
    public partial class GlobalViewUser
    {
        public int idUser { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public Nullable<int> idClient { get; set; }
        public string Email { get; set; }
        public string Password { get; set; }
        public Nullable<System.DateTime> LastUpdated { get; set; }
        public string LastUpdatedBy { get; set; }
        public Nullable<int> Status { get; set; }
        public bool Admin { get; set; }
        public Nullable<bool> isLoggedIn { get; set; }
        public Nullable<System.DateTime> logindatetime { get; set; }
        public Nullable<System.DateTime> logoutdatetime { get; set; }
        public string userimage { get; set; }
    }
}
