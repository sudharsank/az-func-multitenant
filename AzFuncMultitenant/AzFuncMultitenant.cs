using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Configuration;
using Azure.Identity;
using Microsoft.Graph;
using System.Collections.Generic;

namespace AzFuncMultitenant
{
    public class IUser
    {
        public string UserID { get; set; }
        public string DisplayName { get; set; }
        public string Email { get; set; }
    }

    public static class AzFuncMultitenant
    {
        [FunctionName("AzFuncMultitenant")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            List<IUser> userColl = null;

            var clientid = Environment.GetEnvironmentVariable("ClientID");
            var clientSecret = Environment.GetEnvironmentVariable("ClientSecret");

            var tenant = "m365devpractice.onmicrosoft.com";

            if(!string.IsNullOrEmpty(clientid) && !string.IsNullOrEmpty(clientSecret))
            {
                var options = new TokenCredentialOptions
                {
                    AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
                };
                var clientSecretCredential = new ClientSecretCredential(tenant, clientid, clientSecret, options);
                GraphServiceClient client = new GraphServiceClient(clientSecretCredential, new[] { "https://graph.microsoft.com/.default" });
                var users = await client.Users.GetAsync((config) =>
                {
                    // Only request specific properties
                    config.QueryParameters.Select = new[] { "displayName", "id", "mail" };
                    // Get at most 25 results
                    config.QueryParameters.Top = 25;
                    // Sort by display name
                    config.QueryParameters.Orderby = new[] { "displayName" };
                });
                if(users.Value.Count > 0)
                {
                    userColl = new List<IUser>();
                    users.Value.ForEach((user) =>
                    {
                        userColl.Add(new IUser()
                        {
                            UserID = user.Id,
                            DisplayName = user.DisplayName,
                            Email = user.Mail
                        });
                    });
                }
            }          

            return new OkObjectResult(userColl);
        }
    }
}

