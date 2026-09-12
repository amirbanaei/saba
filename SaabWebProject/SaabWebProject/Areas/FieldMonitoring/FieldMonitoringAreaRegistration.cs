using System.Web.Mvc;

namespace SaabWebProject.Areas.FieldMonitoring
{
    public class FieldMonitoringAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "FieldMonitoring";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "FieldMonitoring_default",
                "FieldMonitoring/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}