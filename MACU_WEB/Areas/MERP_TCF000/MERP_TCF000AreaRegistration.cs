using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_TCF000
{
    public class MERP_TCF000AreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "MERP_TCF000";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "MERP_TCF000_default",
                "MERP_TCF000/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}