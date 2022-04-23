using Azure.Identity;
using Microsoft.Graph;
using System;
using System.Configuration;

namespace LinkJar.GraphService
{
    public static class CredentialManager
    {
        public static string ClientId = ConfigurationManager.AppSettings["ClientId"];
        public static string TenantId = ConfigurationManager.AppSettings["TenantId"];
        public static string ClientSecret = ConfigurationManager.AppSettings["ClientSecret"];
        public static string UserName = ConfigurationManager.AppSettings["UserName"];
        public static string UserPassword = ConfigurationManager.AppSettings["UserPassword"];

        public static GraphServiceClient GetGraphClientWithClientSecret()
        {
            // The client credentials flow requires that you request the
            // /.default scope, and preconfigure your permissions on the
            // app registration in Azure. An administrator must grant consent
            // to those permissions beforehand.
            var scopes = new[] { "https://graph.microsoft.com/.default" };

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            var clientSecretCredential = new ClientSecretCredential(TenantId, ClientId, ClientSecret, options);
            var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

            return graphClient;
        }

        public static GraphServiceClient GetGraphClientWithInteractive()
        {
            var options = new InteractiveBrowserCredentialOptions
            {
                TenantId = TenantId,
                ClientId = ClientId,
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
                RedirectUri = new Uri("http://localhost"),
            };

            var interactiveCredential = new InteractiveBrowserCredential(options);
            var graphClient = new GraphServiceClient(interactiveCredential);

            return graphClient;
        }

        public static GraphServiceClient GetGraphClientWithUserIDPwd()
        {
            var scopes = new[] { "User.Read" };

            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            };

            var userNamePasswordCredential = new UsernamePasswordCredential(UserName, UserPassword, TenantId, ClientId, options);
            var graphClient = new GraphServiceClient(userNamePasswordCredential, scopes);

            return graphClient;
        }
    }
}