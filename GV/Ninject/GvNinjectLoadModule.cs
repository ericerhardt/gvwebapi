using System;
using System.Configuration;
using System.Linq;
using System.Web;
using GV.CoFreedomDomain;
using GV.Domain;
using GV.Services;
using Ninject.Activation;
using Ninject.Extensions.NamedScope;
using Ninject.Modules;

namespace GV.Ninject
{
    public class GvNinjectLoadModule : NinjectModule
    {
        public override void Load()
        {
            if (Kernel == null) return;

            Kernel.Bind<ISessionFactoryHelper>().ToMethod(x => new SessionFactoryHelper(StripMetaData(ConfigurationManager.ConnectionStrings["GlobalViewEntities"].ConnectionString))).InSingletonScope();
            Kernel.Bind<ICoFreedomSessionFactory>().ToMethod(x => new CoFreedomSessionFactory(StripMetaData(ConfigurationManager.ConnectionStrings["CoFreedomEntities"].ConnectionString))).InSingletonScope();
            Kernel.Bind<IRepository>().To<Repository>().InScope(HttpRequestOrCall);
            Kernel.Bind<IUnitOfWork>().To<UnitOfWork>().InScope(HttpRequestOrCall);
            Kernel.Bind<ICoFreedomUnitOfWork>().To<CoFreedomUnitOfWork>().InScope(HttpRequestOrCall);
            Kernel.Bind<ICoFreedomRepository>().To<CoFreedomRepository>().InScope(HttpRequestOrCall);
            Kernel.Bind<IEasyLinkFileSaveService>().To<EasyLinkFileSaveService>();
            Kernel.Bind<IEasyLinkService>().To<EasyLinkService>();
            Kernel.Bind<IEasyLinkFileDeleteService>().To<EasyLinkFileDeleteService>();
            Kernel.Bind<IEasyLinkChildManagerService>().To<EasyLinkChildManagerService>();
        }

        private static object HttpRequestOrCall(IContext context)
        {
            return HttpContext.Current ?? CurrentIocResolveScope(context);
        }

        private static object CurrentIocResolveScope(IContext context)
        {
            var parameter = context.Parameters.OfType<NamedScopeParameter>().SingleOrDefault();
            if (parameter != null)
                return parameter.Scope;

            if (context.Request.ParentContext != null)
                return CurrentIocResolveScope(context.Request.ParentContext);

            throw new ArgumentNullException("NamedScopeParameter not found");
        }

        private static string StripMetaData(string connectionString)
        {
            var endIndex = connectionString.IndexOf("data source", StringComparison.OrdinalIgnoreCase);
            return connectionString.Substring(endIndex).Replace("App=EntityFramework\"", string.Empty);
        }
    }
}