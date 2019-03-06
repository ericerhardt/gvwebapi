namespace GVWebapi.Models.Easylink
{
    public class EasyLinkFileSaveModel
    {
        public EasyLinkFileSaveModel(byte[] contents, string fileName)
        {
            Contents = contents;
            FileName = fileName;
        }

        public string FileName { get; }
        public byte[] Contents { get; }

    }
    public class EasyLinkUploadModel
    {
        public string UserName { get; set; }
        public string PeriodDate { get; set; }
        public EasyLinkFileSaveModel EasyLinkFile { get; set; }
    }
}