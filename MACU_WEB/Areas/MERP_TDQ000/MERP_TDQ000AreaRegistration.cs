using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_TDQ000
{
    public class MERP_TDQ000AreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "MERP_TDQ000";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "MERP_TDQ000_default",
                "MERP_TDQ000/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
                
            );
        }
    }
}