using Microsoft.Azure.WebJobs.Description;

namespace SharePointCustomBinding
{
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
        
        public SharePointContextAttribute(string siteUrl, string tenantId, string clientId, string clientSecret)
        {
            TenaneId = tenantId;
            ClientId = clientId;
            ClientSecret = clientSecret;
            SiteUrl = siteUrl;
        }

        public SharePointContextAttribute(string siteUrl, string username, string password)
        {
            SiteUrl = siteUrl;
            Username = username;
            Password = password;
        }
    }
}
