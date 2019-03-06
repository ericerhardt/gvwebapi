namespace GVWebapi.Models.Easylink
{
    public class EasyLinkChildSwitchModel
    {
        public int ExistingCustomerId { get; set; }
        public int ExistingChildId { get; set; }
        public int NewCustomerId { get; set; }
        public int NewChildId { get; set; }
        public bool IsEasyLinkOnly { get; set; }
    }
}