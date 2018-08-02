using System;

namespace GV.CoFreedomDomain.Entities
{
    public class ViewEquipmentAndRate
    {
        public virtual long ViewKey { get; set; }
        public virtual int InvoiceId { get; set; }
        public virtual string EquipmentSerialNumber { get; set; }
        public virtual string EquipmentNumber { get; set; }
        public virtual decimal DifferenceCopies { get; set; }
        public virtual string ContractMeterGroup { get; set; }
        public virtual decimal EffectiveRate { get; set; }
        public virtual DateTime StartDate { get; set; }
        public virtual DateTime EndDate { get; set; }
    }
}