namespace GVWebapi.Controllers
{
    public class EasyLinkChildMatchSaveModel
    {
        public long CustomerId { get; set; }
        public int ChildId { get; set; }
        public bool IsEasyLinkOnly { get; set; }
    }
}