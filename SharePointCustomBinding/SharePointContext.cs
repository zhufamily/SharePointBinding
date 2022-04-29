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
    /// <summary>
    /// Implementation class for custom Sharepoint binding class
    /// </summary>
    [Extension("SharePointContext")]
    public class SharePointContext : IExtensionConfigProvider
    {
        /// <summary>
        /// Initialize the context
        /// </summary>
        /// <param name="context"></param>
        public void Initialize(ExtensionConfigContext context)
        {
            var rule = context.AddBindingRule<SharePointContextAttribute>();
            rule.BindToInput<ClientContext>(BuildItemFromAttribute);
        }

        /// <summary>
        /// Generate and return Client Context for binding 
        /// </summary>
        /// <param name="arg"></param>
        /// <returns></returns>
        private ClientContext BuildItemFromAttribute(SharePointContextAttribute arg)
        {
            try
            {
                if (string.IsNullOrEmpty(arg.Username))
                {
                    if (string.IsNullOrEmpty(arg.SiteUrl) || string.IsNullOrEmpty(arg.ClientId) || string.IsNullOrEmpty(arg.ClientSecret) || string.IsNullOrEmpty(arg.TenaneId))
                        throw new ArgumentException("Missing required parameters for custom Sharepoint binding");

                    // This is another way to get Client Context
                    // AuthenticationManager manager = new AuthenticationManager(arg.ClientId, arg.ClientSecret);
                    // ClientContext context = manager.GetContext(arg.SiteUrl);

                    AuthenticationManager manager = new AuthenticationManager();
                    ClientContext context = manager.GetACSAppOnlyContext(arg.SiteUrl, arg.ClientId, arg.ClientSecret);
                    return context;
                }
                else
                {
                    if (string.IsNullOrEmpty(arg.SiteUrl) || string.IsNullOrEmpty(arg.Username) || string.IsNullOrEmpty(arg.Password))
                        throw new ArgumentException("Missing required parameters for custom Sharepoint binding");


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
