using System.Web.Mvc;

namespace MACU_WEB.Areas.MERP_UAJ000
{
    public class MERP_UAJ000AreaRegistration : AreaRegistration 
    {
        public override string AreaName 
        {
            get 
            {
                return "MERP_UAJ000";
            }
        }

        public override void RegisterArea(AreaRegistrationContext context) 
        {
            context.MapRoute(
                "MERP_UAJ000_default",
                "MERP_UAJ000/{controller}/{action}/{id}",
                new { action = "Index", id = UrlParameter.Optional }
            );
        }
    }
}