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
    
    public partial class EasyLink
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public EasyLink()
        {
            this.EasyLinkItems = new HashSet<EasyLinkItem>();
        }
    
        public long EasyLinkId { get; set; }
        public string FileName { get; set; }
        public string SavedFileName { get; set; }
        public string FileLocation { get; set; }
        public int NumberOfLines { get; set; }
        public System.DateTimeOffset CreatedDateTime { get; set; }
    
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<EasyLinkItem> EasyLinkItems { get; set; }
    }
}
