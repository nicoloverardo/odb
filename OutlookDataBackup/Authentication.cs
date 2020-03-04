using System;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;

namespace OutlookDataBackup
{
    public class AuthenticationHelper
    {
        // The Client ID is used by the application to uniquely identify itself to the v2.0 authentication endpoint.
        private const string ClientId = "";

        private const string BaseUrl = "https://graph.microsoft.com/v1.0";
        private static readonly string[] Scopes = { "Files.ReadWrite.All", "User.ReadBasic.All" };
        private static readonly IPublicClientApplication IdentityClientApp = PublicClientApplicationBuilder.Create(ClientId).Build();
        private static string _tokenForUser;
        private static DateTimeOffset _expiration;
        private static GraphServiceClient _graphClient;

        /// <summary>
        /// Perform login and get the Microsoft Graph client.
        /// </summary>
        /// <returns>The Microsoft Graph client as <see cref="GraphServiceClient"/></returns>
        public static GraphServiceClient GetAuthenticatedClient()
        {
            if (_graphClient != null) return _graphClient;

            // Create Microsoft Graph client.
            _graphClient = new GraphServiceClient(BaseUrl, new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        var token = await GetTokenForUserAsync();
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                    }));

            return _graphClient;
        }

        /// <summary>
        /// Get Token for User.
        /// </summary>
        /// <returns>Token for user.</returns>
        public static async Task<string> GetTokenForUserAsync()
        {
            AuthenticationResult authResult;

            // Get an access token for the given context and resourceId. An attempt is first made to 
            // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
            try
            {
                authResult = await IdentityClientApp.AcquireTokenSilent(Scopes, (await IdentityClientApp.GetAccountsAsync()).FirstOrDefault()).ExecuteAsync().ConfigureAwait(false);
                _tokenForUser = authResult.AccessToken;
            }
            catch (Exception)
            {
                if (_tokenForUser == null || _expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
                {
                    authResult = await IdentityClientApp.AcquireTokenInteractive(Scopes).ExecuteAsync().ConfigureAwait(false);

                    _tokenForUser = authResult.AccessToken;
                    _expiration = authResult.ExpiresOn;
                }
            }

            return _tokenForUser;
        }

        /// <summary>
        /// Sign every user out of the service.
        /// </summary>
        public static async Task SignOut()
        {
            foreach (var user in await IdentityClientApp.GetAccountsAsync())
            {
                await IdentityClientApp.RemoveAsync(user);
            }

            _graphClient = null;
            _tokenForUser = null;
        }

    }
}
