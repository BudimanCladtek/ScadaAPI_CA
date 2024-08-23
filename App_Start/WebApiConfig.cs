using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Cors;

namespace CORSYS_API
{
    public static class WebApiConfig
    {
        public static void Register(HttpConfiguration config)
        {
            // Web API configuration and services

            // Web API routes
            //            var cors = new EnableCorsAttribute("http://192.168.0.200", "*", "*");
            var cors = new EnableCorsAttribute("http://mainca.cor.sys, http://10.112.220.69", "*", "*");
            //var cors = new EnableCorsAttribute("http://192.168.15.100", "*", "*");
            config.EnableCors(cors);
            //config.EnableCors(origins: "http://mainca.cor.sys:80, http://10.112.220.69", headers: "*", methods: "*");
            config.MapHttpAttributeRoutes();

            config.Routes.MapHttpRoute(
                name: "DefaultApiRaw",
                routeTemplate: "api/{controller}/{action}/{raw}",
                defaults: new { controller = "home", action = "index", raw = RouteParameter.Optional }
            );
            config.Routes.MapHttpRoute(
                name: "DefaultApiAct",
                routeTemplate: "api/{controller}/{action}/{id}",
                defaults: new { controller = "home", action = "index", id = RouteParameter.Optional }
            );

            config.Routes.MapHttpRoute(
                name: "DefaultApi",
                routeTemplate: "api/{controller}/{id}",
                defaults: new { id = RouteParameter.Optional }
            );
        }
    }
}
