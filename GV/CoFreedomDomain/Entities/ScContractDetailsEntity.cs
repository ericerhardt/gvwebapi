namespace GV.CoFreedomDomain.Entities
{
    public class ScContractDetailsEntity
    {
        public virtual int ContractDetailId { get; set; }
        public virtual ScContractsEntity Contract { get; set; }
        public virtual ScEquipmentEntity Equipment { get; set; }
    }
}