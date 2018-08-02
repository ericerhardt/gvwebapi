using GV.Configuration;

namespace GVWebapi.Configuration
{
    public class GlobalViewConfiguration : IGlobalViewConfiguration
    {
        public GlobalViewConfiguration(string globalViewConnectionString, string easyLinkFileSavePath)
        {
            GlobalViewConnectionString = globalViewConnectionString;
            EasyLinkFileSavePath = easyLinkFileSavePath;
        }

        public string GlobalViewConnectionString { get; }
        public string EasyLinkFileSavePath { get; }
    }
}