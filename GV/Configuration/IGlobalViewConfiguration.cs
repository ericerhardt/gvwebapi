namespace GV.Configuration
{
    public interface IGlobalViewConfiguration
    {
        string GlobalViewConnectionString { get; }
        string EasyLinkFileSavePath { get; }
    }
}