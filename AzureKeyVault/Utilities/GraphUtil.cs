using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace AzureKeyVault.Utilities;

public static class GraphUtil
{
    /// <summary>
    /// Get the Azure Token
    /// </summary>
    /// <param name="tenantId">Azure TenantId</param>
    /// <param name="clientId">Application/ClientId</param>
    /// <param name="clientSecret">Application/Client Secret</param>
    /// <param name="version">Options Version, if omitted defaults to v2.0</param>
    /// <returns></returns>
    public static async Task<AuthenticationResult> GetToken(this string tenantId, string clientId, string clientSecret, string version = "v2.0")
    {
        string[] scope = {".default"};
        string authority = $"https://login.microsoftonline.com/{tenantId}/{version}";
        
        IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(clientId)
            .WithClientSecret(clientSecret)
            .WithAuthority(new Uri(authority))
            .Build();

        return await app.AcquireTokenForClient(scope)
            .ExecuteAsync();
    }

    public static GraphServiceClient ClientSecret(this string tenantId, string clientId, string clientSecret, string[]? scopes = null)
    {
        scopes ??= new[] {".default"};
        
        // using Azure.Identity;
        TokenCredentialOptions options = new ()
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
        };

        // https://learn.microsoft.com/dotnet/api/azure.identity.clientsecretcredential
        ClientSecretCredential clientSecretCredential = new(
            tenantId, clientId, clientSecret, options);

        GraphServiceClient graphClient = new (clientSecretCredential, scopes);
        return graphClient;
    }
    
    
    public static GraphServiceClient IntegratedWindowsProvider(this string tenantId, string clientId)
    {
        string[] scopes = new[] { "User.Read" };
        IPublicClientApplication? pca = PublicClientApplicationBuilder
            .Create(clientId)
            .WithTenantId(tenantId)
            .Build();

        // DelegateAuthenticationProvider is a simple auth provider implementation
        // that allows you to define an async function to retrieve a token
        // Alternatively, you can create a class that implements IAuthenticationProvider
        // for more complex scenarios
        DelegateAuthenticationProvider authProvider = new(async (request) => {
            // Use Microsoft.Identity.Client to retrieve token
            AuthenticationResult? result = await pca.AcquireTokenByIntegratedWindowsAuth(scopes).ExecuteAsync();

            request.Headers.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", result.AccessToken);
        });

        GraphServiceClient graphClient = new(authProvider);
        return graphClient;
    }
}