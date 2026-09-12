using System.Web.Mvc;

namespace SaabWebProject.Areas.ManagementAccounting
{
    public class ManagementAccountingAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "ManagementAccounting";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "ManagementAccounting_default",
                "ManagementAccounting/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}