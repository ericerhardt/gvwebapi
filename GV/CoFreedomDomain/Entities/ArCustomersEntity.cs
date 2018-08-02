namespace GV.CoFreedomDomain.Entities
{
    public class ArCustomersEntity
    {
        public virtual int CustomerId { get; set; }
        public virtual string CustomerName { get; set; }
        public virtual string Address { get; set; }
        public virtual string City { get; set; }
        public virtual string State { get; set; }
        public virtual string Zip { get; set; }
        
    }
}