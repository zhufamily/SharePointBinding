using Microsoft.Azure.WebJobs.Description;

namespace SharePointCustomBinding
{
    /// <summary>
    /// Custom Sharepoint binding attribute class
    /// </summary>
    [Binding]
    [AttributeUsage(AttributeTargets.Parameter | AttributeTargets.ReturnValue)]
    public class SharePointContextAttribute : Attribute
    {
        [AutoResolve]
        public string? TenaneId { get; set; }
        [AutoResolve]
        public string? ClientId { get; set; }
        [AutoResolve]
        public string? ClientSecret { get; set; }
        [AutoResolve]
        public string? SiteUrl { get; set; }
        [AutoResolve]
        public string? Username { get; set; }
        [AutoResolve]
        public string? Password { get; set; }
        
        /// <summary>
        /// Constructor for client secret
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="tenantId"></param>
        /// <param name="clientId"></param>
        /// <param name="clientSecret"></param>
        public SharePointContextAttribute(string siteUrl, string tenantId, string clientId, string clientSecret)
        {
            TenaneId = tenantId;
            ClientId = clientId;
            ClientSecret = clientSecret;
            SiteUrl = siteUrl;
        }

        /// <summary>
        /// Constructor for username and password
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="username"></param>
        /// <param name="password"></param>
        public SharePointContextAttribute(string siteUrl, string username, string password)
        {
            SiteUrl = siteUrl;
            Username = username;
            Password = password;
        }
    }
}
