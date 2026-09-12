using System.Web.Mvc;

namespace SaabWebProject.Areas.AlarmsAndLog
{
    public class AlarmsAndLogAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "AlarmsAndLog";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "AlarmsAndLog_default",
                "AlarmsAndLog/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}