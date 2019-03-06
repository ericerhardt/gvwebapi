namespace GVWebapi.Models.Easylink
{
    public class EasyLinkChildMatchSaveModel
    {
        public int CustomerId { get; set; }
        public int ChildId { get; set; }
        public bool IsEasyLinkOnly { get; set; }
    }
}