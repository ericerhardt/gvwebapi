namespace GV.CoFreedomDomain.Entities
{
    public class ScEquipmentCustomPropertiesEntity
    {
        public virtual int EquipmentCustomPropertyId { get; set; }
        public virtual ScEquipmentEntity Equipment { get; set; }
        public virtual int ShAttributeId { get; set; }
        public virtual int? IdVal { get; set; }
        public virtual string TextVal { get; set; }
    }
}