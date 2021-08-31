using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_UAK000
{
    public class MERP_UAK000AreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "MERP_UAK000";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "MERP_UAK000_default",
                "MERP_UAK000/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}