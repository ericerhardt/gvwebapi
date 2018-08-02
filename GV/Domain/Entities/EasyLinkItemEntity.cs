using System;

namespace GV.Domain.Entities
{
    public class EasyLinkItemEntity
    {
        public virtual long EasyLinkItemId { get; set; }
        public virtual EasyLinkEntity EasyLink { get; set; }
        public virtual int Child { get; set; }
        public virtual string Email { get; set; }
        public virtual string Fax { get; set; }
        public virtual DateTime DateTime { get; set; }
        public virtual string Description { get; set; }
        public virtual decimal Duration { get; set; }
        public virtual int Pages { get; set; }
        public virtual decimal Charge { get; set; }
        public virtual string MessageNumber { get; set; }
    }
}