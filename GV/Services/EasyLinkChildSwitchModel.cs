namespace GV.Services
{
    public class EasyLinkChildSwitchModel
    {
        public long ExistingCustomerId { get; set; }
        public int ExistingChildId { get; set; }
        public long NewCustomerId { get; set; }
        public int NewChildId { get; set; }
        public bool IsEasyLinkOnly { get; set; }
    }
}