using GVWebapi.Helpers;
using System.Web.Http;

using System.Web.Http.Cors;

namespace GVWebapi
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            var enableCorsAttribute = new  EnableCorsAttribute("*","*","*");
            config.EnableCors(enableCorsAttribute);
            config.MapHttpAttributeRoutes();
           // config.Filters.Add(new GlobalViewAuthorizeAttribute());
            config.Routes.MapHttpRoute("DefaultApi","api/{controller}/{id}",new { id = RouteParameter.Optional });
        }
    }
}
