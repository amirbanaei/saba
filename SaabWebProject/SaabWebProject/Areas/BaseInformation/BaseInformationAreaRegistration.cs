using System.Web.Mvc;

namespace SaabWebProject.Areas.BaseInformation
{
    public class BaseInformationAreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "BaseInformation";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "BaseInformation_default",
                "BaseInformation/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}