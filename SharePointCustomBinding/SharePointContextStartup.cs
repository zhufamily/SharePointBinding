
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Hosting;
using SharePointCustomBinding;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[assembly: WebJobsStartup(typeof(SharePointContextStartup))]
namespace SharePointCustomBinding
{
    /// <summary>
    /// Custom binding Web Job initializer
    /// </summary>
    public class SharePointContextStartup : IWebJobsStartup
    {
        /// <summary>
        /// Add custom binding extensions to Web Job
        /// </summary>
        /// <param name="builder"></param>
        /// <exception cref="ArgumentNullException"></exception>
        public void Configure(IWebJobsBuilder builder)
        {
            if (builder == null)
            {
                throw new ArgumentNullException(nameof(builder));
            }

            builder.AddExtension<SharePointContext>();
        }
    }
}

