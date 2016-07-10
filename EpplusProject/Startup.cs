using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(EpplusProject.Startup))]
namespace EpplusProject
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
