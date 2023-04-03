using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using Azure.Identity;
using Microsoft.Graph.Models;
using Microsoft.Graph.Models.ODataErrors;
using System.Text;
using System.Linq;
using System.Collections.Generic;

namespace GraphApiSubscription
{
    public static class Function1
    {

        [FunctionName("Create")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            var graphClient = GetGraphServiceClient();
            var list = new List<Subscription>();
            list.Add(new Subscription
            {
                ChangeType = "created",
                //NotificationUrl = "https://webhook.azurewebsites.net/api/send/myNotifyClient",
                NotificationUrl = "<url>",
                Resource = "/chats/getAllMembers",
                //Resource = "/teams/getAllMessages",
                ExpirationDateTime = DateTime.Now.AddHours(1),
                ClientState = "secretClientValue",
                LatestSupportedTlsVersion = "v1_2",
                EncryptionCertificate = "<公開キー>",
                EncryptionCertificateId = "<EncryptionCertificateId>",
                IncludeResourceData = true,
            });
            list.Add(new Subscription
            {
                ChangeType = "created",
                //NotificationUrl = "https://webhook.azurewebsites.net/api/send/myNotifyClient",
                NotificationUrl = "<url>",
                Resource = "/teams/getAllMessages",
                ExpirationDateTime = DateTime.Now.AddHours(1),
                ClientState = "secretClientValue",
                LatestSupportedTlsVersion = "v1_2",
                EncryptionCertificate = "<公開キー>",
                EncryptionCertificateId = "<EncryptionCertificateId>",
                IncludeResourceData = true,
            });


            StringBuilder sb = new StringBuilder();
            foreach (var requestBody in list)
            {
                try
                {
                    var result = await graphClient.Subscriptions.PostAsync(requestBody);
                    sb.AppendLine($"{nameof(result.Id)}:{result.Id}");
                    sb.AppendLine($"{nameof(result.ApplicationId)}:{result.ApplicationId}");
                    sb.AppendLine($"{nameof(result.ExpirationDateTime)}:{result.ExpirationDateTime}");
                    sb.AppendLine($"{nameof(result.NotificationUrl)}:{result.NotificationUrl}");
                    sb.AppendLine();
                }
                catch (ODataError e)
                {
                    log.LogInformation(e.ToString());
                    log.LogInformation(e.Error.Message);
                    sb.AppendLine(e.ToString() + Environment.NewLine + e.Error.Message);
                    sb.AppendLine();
                    continue;
                }
                catch (Exception e)
                {
                    log.LogInformation(e.ToString());
                    sb.AppendLine(e.ToString());
                    sb.AppendLine();
                    continue;
                }
            }

            return new OkObjectResult(sb.ToString());
        }

        [FunctionName("List")]
        public static async Task<IActionResult> List(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");

            var graphClient = GetGraphServiceClient();

            var responseMessage = string.Empty;
            try
            {
                var result = await graphClient.Subscriptions.GetAsync();

                StringBuilder sb = new StringBuilder();
                if(result.Value.Count == 0)
                {
                    sb.AppendLine("Empty");
                }
                foreach(var subscription in result.Value)
                {
                    sb.AppendLine($"{nameof(subscription.Id)}:{subscription.Id}");
                    sb.AppendLine($"{nameof(subscription.ApplicationId)}:{subscription.ApplicationId}");
                    sb.AppendLine($"{nameof(subscription.NotificationUrl)}:{subscription.NotificationUrl}");
                    sb.AppendLine($"{nameof(subscription.ChangeType)}:{subscription.ChangeType}");
                    sb.AppendLine($"{nameof(subscription.Resource)}:{subscription.Resource}");
                    sb.AppendLine();
                }
                
                responseMessage = sb.ToString();
            }
            catch (ODataError e)
            {
                log.LogInformation(e.ToString());
                log.LogInformation(e.Error.Message);
            }
            catch (Exception e)
            {
                log.LogInformation(e.ToString());
            }

            return new OkObjectResult(responseMessage);
        }

        private static GraphServiceClient GetGraphServiceClient()
        {
            // The client credentials flow requires that you request the
            // /.default scope, and preconfigure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            // Multi-tenant apps can use "common",
            // single-tenant apps must use the tenant ID from the Azure portal
            var tenantId = "<tenantId>";

            // Values from app registration
            var clientId = "<clientId>";
            var clientSecret = "<clientSecret>";

            // using Azure.Identity;
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
            var clientSecretCredential = new ClientSecretCredential(
                tenantId, clientId, clientSecret, options);

            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);
            return graphClient;
        }
    }
}
