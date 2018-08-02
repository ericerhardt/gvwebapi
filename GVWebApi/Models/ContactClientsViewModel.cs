namespace GVWebapi.Models
{
    public class ContactClientsViewModel
    {
        public ContactClientsViewModel(long id, long fprId, int customerId, string name, bool isChecked)
        {
            ContactClientId = id;
            FprContactId = fprId;
            CustomerId = customerId;
            CustomerName = name;
            Selected = isChecked;
        }

        public long ContactClientId { get; set; }
        public long FprContactId { get; set; }
        public int CustomerId { get; set; }
        public string CustomerName { get; set; }
        public bool Selected { get; set; }
    }
}