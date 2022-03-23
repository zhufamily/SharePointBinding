using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs.Description;
using Microsoft.Azure.WebJobs.Host.Config;
using Microsoft.SharePoint.Client;
using PnP.Core.Services;
using PnP.Framework;

namespace SharePointCustomBinding
{
    [Extension("SharePointContext")]
    public class SharePointContext : IExtensionConfigProvider
    {
        public void Initialize(ExtensionConfigContext context)
        {
            var rule = context.AddBindingRule<SharePointContextAttribute>();
            rule.BindToInput<ClientContext>(BuildItemFromAttribute);
        }

        private ClientContext BuildItemFromAttribute(SharePointContextAttribute arg)
        {
            try
            {
                if (string.IsNullOrEmpty(arg.Username))
                {
                    // AuthenticationManager manager = new AuthenticationManager(arg.ClientId, arg.ClientSecret);
                    // ClientContext context = manager.GetContext(arg.SiteUrl);

                    AuthenticationManager manager = new AuthenticationManager();
                    ClientContext context = manager.GetACSAppOnlyContext(arg.SiteUrl, arg.ClientId, arg.ClientSecret);
                    return context;
                }
                else
                {
                    SecureString encpwd = new SecureString();
                    foreach (char c in arg.Password)
                    {
                        encpwd.AppendChar(c);
                    }
                    AuthenticationManager manager = new AuthenticationManager(arg.Username, encpwd);
                    ClientContext context = manager.GetContext(arg.SiteUrl);
                    return context;
                }
            } 
            catch (Exception)
            {
                throw;
            }
        }
    }
}
