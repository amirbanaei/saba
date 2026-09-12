using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;

namespace SaabWebProject
{ 
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");
            routes.MapRoute(
                "Area",
                "",
                new { area = "Users", controller = "Login", action = "Login" }
                );
            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );
            routes.MapRoute(
                name: "Contracts",
                url: "Contracts/{controller}/{action}/{id}",
                defaults: new { area = "Contracts", controller = "Contract", action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}
