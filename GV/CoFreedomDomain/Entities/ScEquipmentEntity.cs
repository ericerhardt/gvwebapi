using System.Collections.Generic;

namespace GV.CoFreedomDomain.Entities
{
    public class ScEquipmentEntity
    {
        private readonly IList<ScEquipmentCustomPropertiesEntity> _customProperties;
        private readonly IList<ScContractDetailsEntity> _contractDetails;

        public ScEquipmentEntity()
        {
            _customProperties = new List<ScEquipmentCustomPropertiesEntity>();
            _contractDetails = new List<ScContractDetailsEntity>();
        }

        public virtual int EquipmentId { get; set; }
        public virtual string EquipmentNumber { get; set; }
        public virtual string SerialNumber { get; set; }
        public virtual ArCustomersEntity Location { get; set; }
        public virtual int CustomerId { get; set; }
   
        public virtual bool Active { get; set; }
        public virtual IcModelEntity Model { get; set; }

        public virtual IEnumerable<ScContractDetailsEntity> ContractDetails => _contractDetails;
        public virtual IEnumerable<ScEquipmentCustomPropertiesEntity> CustomProperties => _customProperties;
    }
}