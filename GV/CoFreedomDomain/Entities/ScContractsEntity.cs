using System.Collections.Generic;

namespace GV.CoFreedomDomain.Entities
{
    public class ScContractsEntity
    {
        private readonly IList<ScContractMeterGroupsEntity> _contractMeterGroups;

        public ScContractsEntity()
        {
            _contractMeterGroups = new List<ScContractMeterGroupsEntity>();
        }

        public virtual int ContractId { get; set; }
        public virtual int CustomerId { get; set; }

        public virtual IEnumerable<ScContractMeterGroupsEntity> ContractMeterGroups => _contractMeterGroups;
    }
}