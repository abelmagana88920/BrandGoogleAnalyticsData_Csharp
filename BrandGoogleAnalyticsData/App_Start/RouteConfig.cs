using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Routing;
using System.Web.Http;
using System.Web.Http.WebHost;

namespace BrandGoogleAnalyticsData
{
    public class RouteConfig
    {
        public static void RegisterRoutes(RouteCollection routes)
        {
            routes.IgnoreRoute("{resource}.axd/{*pathInfo}");

            routes.MapRoute(
                name: "GetCustomerName",
                url: "Customer/Name/{customer_name}",
                defaults: new { controller = "Customer", action = "Name", customer_name = UrlParameter.Optional }
            );

            routes.MapRoute(
                name: "GetCustomerNumber",
                url: "Customer/Number/{customer_no}",
                defaults: new { controller = "Customer", action = "Number", customer_no = UrlParameter.Optional }
            );

            routes.MapRoute(
               name: "GetCustomerEmail",
               url: "Customer/Emails/{customer_no}",
               defaults: new { controller = "Customer", action = "Emails", customer_no = UrlParameter.Optional }
           );

            routes.MapRoute(
               name: "GetCustomeFTP",
               url: "Customer/FTPs/{customer_no}",
               defaults: new { controller = "Customer", action = "FTPs", customer_no = UrlParameter.Optional }
           ); 


           routes.MapRoute(
              name: "SendRequest",
              url: "Customer/SendRequest/{customer_no}",
              defaults: new { controller = "Customer", action = "SendRequest", customer_no = UrlParameter.Optional }
          );

           routes.MapRoute(
           name: "FeedEdit",
           url: "Customer/FeedEdit/{if_id}",
           defaults: new { controller = "Customer", action = "FeedEdit", if_id = UrlParameter.Optional }
       );

          

            routes.MapRoute(
                name: "Default",
                url: "{controller}/{action}/{id}",
                defaults: new { controller = "Home", action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}