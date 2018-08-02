using System;

namespace GV.Domain.Entities
{
    public class EasyLinkChildMatchEntity
    {
        public virtual long EasyLinkChildMatchId { get; set; }
        public virtual long CustomerId { get; set; }
        public virtual int ChildId { get; set; }
        public virtual bool IsEasyLinkOnly { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }
        public virtual DateTimeOffset? ModifiedDateTime { get; set; }
        public virtual bool IsDeleted { get; set; }
    }
}