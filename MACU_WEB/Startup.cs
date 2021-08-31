using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(MACU_WEB.Startup))]
namespace MACU_WEB
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
