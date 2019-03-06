namespace GVWebapi.Models.Easylink
{
    public class EasyLinkChildManagerModel
    {
        public long EasyLinkChildMatchId { get; set; }
        public string CustomerName { get; set; }
        public long CustomerId { get; set; }
        public int ChildId { get; set; }
        public bool IsEasyLinkOnly { get; set; }
    }
}