# SharePoint Online Binding
## _A custom binding for SharePoint Online in Azure function_
## Features

- Input / output binding for SharePoint Online Client Context 
## Dependencies
- .Net 6.x
- Microsoft.Azure.WebJob 3.0.33+
- Microsoft.Azure.WebJob.Core 3.0.33+
- PnP.Framework 1.10.0+

## Installation

It is available from public Nuget source

Install the package and all dependencies.

```sh
PM> Install-Package SharePointCustomBinding -version <latest_version>
```

## Samples

SharePoint Online Binding can be used as in following samples with user name and password

```sh
[FunctionName("<your_function_name>")]
public static string <your_function_name>(<your_function_trigger>,
[SharePointContext("<your_sp_online_url>", "<your_username>", "<your_password>")] ClientContext context,
ILogger log)
{
    log.LogInformation($"Function triggered");
    using(context) {
        // your actions with SharePoint Online 
    }
    return "done!";
}
```

If you prefer application / client id plus client secret, then see the following sample 

```sh
[FunctionName("<your_function_name>")]
public static string <your_function_name>(<your_function_trigger>,
[SharePointContext("<your_sp_online_url>", "<your_tenantid>", "<your_clientid>", "<your_clientsecret>")] ClientContext context,
ILogger log)
{
    log.LogInformation($"Function triggered");
    using(context) {
        // your actions with SharePoint Online 
    }
    return "done!";
}
```

For more about how to use SharePoint client context, please refer to https://www.nuget.org/packages/Microsoft.SharePointOnline.CSOM/. 

## License

Free software, absolutely no warranty, use at your own risk!
