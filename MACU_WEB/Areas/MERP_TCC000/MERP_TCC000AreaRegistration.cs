using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_TCC000
{
    public class MERP_TCC000AreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "MERP_TCC000";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "MERP_TCC000_default",
                "MERP_TCC000/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}