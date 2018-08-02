namespace GV.Services
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
}