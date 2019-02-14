using System;
using System.Collections.Generic;

namespace GV.Domain.Entities
{
    public class EasyLinkEntity
    {
        private readonly IList<EasyLinkItemEntity> _items = new List<EasyLinkItemEntity>();

        public virtual long EasyLinkId { get; set; }
        public virtual string FileName { get; set; }
        public virtual string FileLocation { get; set; }
        public virtual string SavedFileName { get; set; }
        public virtual int NumberOfLines { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }

        public virtual IEnumerable<EasyLinkItemEntity> Items => _items;
    }
    public class EasyLinkImportHistoryEntity
    {
        private readonly IList<EasyLinkDataEntity> _items = new List<EasyLinkDataEntity>();

        public virtual int ImportID { get; set; }
        public virtual string FileName { get; set; }
        public virtual string FileLocation { get; set; }
        public virtual DateTime PeriodDate { get; set; } 
        public virtual int ImportedRecords { get; set; }
        public virtual DateTimeOffset CreatedDateTime { get; set; }

        public virtual IEnumerable<EasyLinkDataEntity> Items => _items;
    }
}