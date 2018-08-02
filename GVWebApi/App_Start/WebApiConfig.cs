using System.Web.Http;
using System.Web.Http.Cors;
namespace GVWebapi
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            var enableCorsAttribute = new  EnableCorsAttribute("*","Origin, Content-Type, Accept","GET, PUT, POST, DELETE, OPTIONS");
            config.EnableCors(enableCorsAttribute);
            config.MapHttpAttributeRoutes();
            config.Routes.MapHttpRoute("DefaultApi","api/{controller}/{id}",new { id = RouteParameter.Optional });
        }
    }
}
