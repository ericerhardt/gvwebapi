using System;
using System.Configuration;
using System.Linq;
using System.Web;
using GV.CoFreedomDomain;
using GV.Configuration;
using GV.Domain;
using GV.Ninject;
using GVWebapi.Configuration;
using GVWebapi.Services;
using Ninject;
using Ninject.Activation;
using Ninject.Extensions.Conventions;
using Ninject.Extensions.NamedScope;
using Ninject.Modules;

namespace GVWebapi.Ninject
{
    public class NinjectApiLoadModule : NinjectModule
    {
        public override void Load()
        {
            if (Kernel == null) return;

            Kernel.Bind<IGlobalViewConfiguration>().ToMethod(context => new GlobalViewConfiguration(
                StripMetaData(ConfigurationManager.ConnectionStrings["GlobalViewEntities"].ConnectionString),
                ConfigurationManager.AppSettings["EasyLink.FileSavePath"]
            ));

            Kernel.Bind<ICoFreedomDeviceService>().To<CoFreedomDeviceService>();
            Kernel.Bind<ICycleHistoryService>().To<CycleHistoryService>();
            Kernel.Bind<ICyclePeriodService>().To<CyclePeriodService>();
            Kernel.Bind<IDeviceService>().To<DeviceService>();
            Kernel.Bind<ILocationsService>().To<LocationsService>();
            Kernel.Bind<IEditScheduleService>().To<EditScheduleService>();
            Kernel.Bind<IScheduleService>().To<ScheduleService>();
            Kernel.Bind<IReconciliationService>().To<ReconciliationService>();
            Kernel.Bind<IScheduleServicesService>().To<ScheduleServicesService>();
        }

        private static string StripMetaData(string connectionString)
        {
            var endIndex = connectionString.IndexOf("data source", StringComparison.OrdinalIgnoreCase);
            return connectionString.Substring(endIndex).Replace("App=EntityFramework\"", string.Empty);
        }
    }
}