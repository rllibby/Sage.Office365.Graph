/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Sage.Office365.Graph.Authentication.Common;
using Sage.Office365.Graph.Authentication.Interfaces;
using Sage.WebAuthenticationBroker;
using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Sage.Office365.Graph.Authentication
{
    /// <summary>
    /// Authentication provider for Azure AD /Graph.
    /// </summary>
    [SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable")]
    public class AuthenticationProvider : IAuthenticationProvider2
    {
        #region Private constants

        private const int AuthWindowHeight = 677;

        #endregion

        #region Private fields

        private readonly IList<string> _scopes = new List<string>( );
        private readonly IHttpProvider _httpProvider;
        private IAuthenticationStore _store;
        private DateTimeOffset _expiration;
        private readonly string _tokenServiceUrl;
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _tenantId;
        private readonly string _redirectUri;
        private string _accessToken;
        private string _refreshToken;
        private bool _appBased;

        #endregion

        #region Private methods

        /// <summary>
        /// Throws the service error as a ServiceException.
        /// </summary>
        /// <param name="code">The graph error code.</param>
        /// <param name="errorDescription">The error description.</param>
        private static void ThrowServiceError(GraphErrorCode code, string errorDescription)
        {
            throw new ServiceException(new Error { Code = code.ToString(), Message = errorDescription });
        }

        /// <summary>
        /// Throws an exception if an error is found in the dictionary of returned values.
        /// </summary>
        /// <param name="responseValues">The oAuth response values.</param>
        private void ThrowIfError(IDictionary<string, string> responseValues)
        {
            if (responseValues == null) return;

            string error = null;
            string errorDescription = null;

            if (responseValues.TryGetValue(Params.ErrorDescriptionKeyName, out errorDescription) || responseValues.TryGetValue(Params.ErrorKeyName, out error))
            {
                ThrowServiceError(GraphErrorCode.AuthenticationFailure, errorDescription ?? error);
            }
        }

        /// <summary>
        /// Signs the request headers with the authorization bearer token.
        /// </summary>
        /// <param name="request">The request to sign.</param>
        /// <returns>True if the access token was valid and the request was signed, otherwise false.</returns>
        private bool SignRequest(HttpRequestMessage request)
        {
            if (string.IsNullOrEmpty(_accessToken)) return false;
            if (_expiration <= DateTimeOffset.Now.UtcDateTime.AddMinutes(5)) return false;

            request.Headers.Authorization = new AuthenticationHeaderValue(Headers.Bearer, _accessToken);

            return true;
        }

        /// <summary>
        /// Refresh the current access token, if possible.
        /// </summary>
        /// <returns>The task to await.</returns>
        private Task RefreshAccessTokenAsync()
        {
            if (!string.IsNullOrEmpty(_refreshToken))
            {
                return SendTokenRequestAsync(GetRefreshTokenRequestBody(_refreshToken));
            }

            return Task.FromResult(0);
        }

        /// <summary>
        /// Sends the token request which can either be app or user based.
        /// </summary>
        /// <param name="requestBodyString">The body for the request.</param>
        /// <returns>The task to wait on.</returns>
        private async Task SendTokenRequestAsync(string requestBodyString)
        {
            var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, _tokenServiceUrl);

            httpRequestMessage.Content = new StringContent(requestBodyString, Encoding.UTF8, Headers.FormUrlEncodedContentType);

            using (var authResponse = await _httpProvider.SendAsync(httpRequestMessage))
            {
                using (var responseStream = await authResponse.Content.ReadAsStreamAsync())
                {
                    var responseValues = _httpProvider.Serializer.DeserializeObject<IDictionary<string, string>>(responseStream);

                    if (responseValues != null)
                    {
                        ThrowIfError(responseValues);

                        _accessToken = responseValues[Params.AccessTokenKeyName];
                        _expiration = DateTimeOffset.UtcNow.Add(new TimeSpan(0, 0, int.Parse(responseValues[Params.ExpiresInKeyName])));

                        if (responseValues.ContainsKey(Params.RefreshTokenKeyName))
                        { 
                            _refreshToken = responseValues[Params.RefreshTokenKeyName];

                            if (_store != null) _store.RefreshToken = _refreshToken;
                        }

                        return;
                    }

                    ThrowServiceError(GraphErrorCode.AuthenticationFailure, ServiceMessages.NoTokenReturned);
                }
            }
        }

        /// <summary>
        /// Gets the authentication result.
        /// </summary>
        /// <returns>The task to wait on.</returns>
        private async Task GetAuthenticationResultAsync()
        {
            if (_appBased)
            { 
                await SendTokenRequestAsync(GetAccessTokenRequestBody());

                return;
            }

            var code = GetAuthorizationCodeAsync(_redirectUri);

            if (!string.IsNullOrEmpty(code)) await SendTokenRequestAsync(GetCodeRedemptionRequestBody(code));
        }

        /// <summary>
        /// Gets the authorization code.
        /// </summary>
        /// <param name="redirectUri">The redirect uri for the oAuth request.</param>
        /// <returns>The oAuth code on success, null on failure.</returns>
        private string GetAuthorizationCodeAsync(string redirectUri)
        {
            WebAuthenticationBroker.WebAuthenticationBroker.Height = AuthWindowHeight;

            var requestUri = new Uri(GetAuthorizationCodeRequestUrl(redirectUri));
            var result = WebAuthenticationBroker.WebAuthenticationBroker.Authenticate(WebAuthenticationOptions.None, requestUri, new Uri(redirectUri));

            IDictionary<string, string> authenticationResponseValues = null;

            if ((result != null) && !string.IsNullOrEmpty(result.ResponseData))
            {
                authenticationResponseValues = UrlHelper.GetQueryOptions(new Uri(result.ResponseData));

                ThrowIfError(authenticationResponseValues);
            }
            else if ((result != null) && (result.ResponseStatus == WebAuthenticationStatus.UserCancel))
            {
                ThrowServiceError(GraphErrorCode.AuthenticationFailure, ServiceMessages.UserCancel);
            }

            var code = string.Empty;

            return ((authenticationResponseValues != null) && authenticationResponseValues.TryGetValue(Params.CodeKeyName, out code)) ? code : null;
        }

        /// <summary>
        /// Gets the request body for generating an application based access token.
        /// </summary>
        /// <returns>The request body for the call.</returns>
        private string GetAccessTokenRequestBody()
        {
            var requestBodyStringBuilder = new StringBuilder();

            requestBodyStringBuilder.AppendFormat("{0}={1}", Params.ClientIdKeyName, _clientId);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.ClientSecretKeyName, _clientSecret);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.GrantTypeKeyName, Params.ClientCredentions);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", _scopes)));

            return requestBodyStringBuilder.ToString();
        }

        /// <summary>
        /// Gets the administration consent url.
        /// </summary>
        /// <param name="returnUrl">The redirect uri for the request.</param>
        /// <returns>The oAuth2 request url.</returns>
        private string GetAdminConsentRequestUrl(string redirectUri)
        {
            var requestUriStringBuilder = new StringBuilder();

            requestUriStringBuilder.Append(string.Format(Endpoints.AdminConsentServiceUrl, _tenantId));
            requestUriStringBuilder.AppendFormat("?{0}={1}", Params.ClientIdKeyName, _clientId);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Params.StateKeyName, Params.ConsentTypeValueName);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Params.RedirectUriKeyName, WebUtility.UrlEncode(redirectUri));

            return requestUriStringBuilder.ToString();
        }

        /// <summary>
        /// Gets the logout url.
        /// </summary>
        /// <param name="returnUrl">The redirect uri for the request.</param>
        /// <returns>The oAuth2 request url.</returns>
        private string GetLogoutRequestUrl(string redirectUri)
        {
            var requestUriStringBuilder = new StringBuilder();

            requestUriStringBuilder.Append(Endpoints.LogoutServiceUrl);
            requestUriStringBuilder.AppendFormat("?{0}={1}", Params.ClientIdKeyName, _clientId);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Params.PostRedirectUriKeyName, redirectUri);

            return requestUriStringBuilder.ToString();
        }

        /// <summary>
        /// Gets the request url for the oAuth2 authentication using the code flow.
        /// </summary>
        /// <param name="redirectUri">The redirect uri for the request.</param>
        /// <returns>The oAuth2 authentication url.</returns>
        private string GetAuthorizationCodeRequestUrl(string redirectUri)
        {
            var requestUriStringBuilder = new StringBuilder();

            requestUriStringBuilder.Append(Endpoints.AuthenticationServiceUrl);
            requestUriStringBuilder.AppendFormat("?{0}={1}", Params.RedirectUriKeyName, WebUtility.UrlEncode(redirectUri));
            requestUriStringBuilder.AppendFormat("&{0}={1}", Params.ClientIdKeyName, _clientId);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Params.ResponseTypeKeyName, Params.CodeKeyName);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Params.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", _scopes)));

            return requestUriStringBuilder.ToString();
        }

        /// <summary>
        /// Gets the request body for redeeming an authorization code for an access token.
        /// </summary>
        /// <param name="code">The authorization code to redeem.</param>
        /// <param name="returnUrl">The return URL for the request. Defaults to the service info value.</param>
        /// <returns>The request body for the code redemption call.</returns>
        private string GetCodeRedemptionRequestBody(string code)
        {
            var requestBodyStringBuilder = new StringBuilder();

            requestBodyStringBuilder.AppendFormat("{0}={1}", Params.RedirectUriKeyName, _redirectUri);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.ClientIdKeyName, _clientId);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.CodeKeyName, code);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.GrantTypeKeyName, Params.AuthorizationCodeGrantType);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", _scopes)));

            return requestBodyStringBuilder.ToString();
        }

        /// <summary>
        /// Gets the request body for redeeming a refresh token for an access token.
        /// </summary>
        /// <param name="refreshToken">The refresh token to redeem.</param>
        /// <returns>The request body for the redemption call.</returns>
        private string GetRefreshTokenRequestBody(string refreshToken)
        {
            var requestBodyStringBuilder = new StringBuilder();

            requestBodyStringBuilder.AppendFormat("{0}={1}", Params.RedirectUriKeyName, _redirectUri);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.ClientIdKeyName, _clientId);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.RefreshTokenKeyName, refreshToken);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.GrantTypeKeyName, Params.RefreshTokenKeyName);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Params.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", _scopes)));

            return requestBodyStringBuilder.ToString();
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor for user based authentication.
        /// </summary>
        /// <param name="clientId">The client id for the application.</param>
        /// <param name="store">The optional authentication store to use for the refresh token.</param>
        public AuthenticationProvider(string clientId, IAuthenticationStore tokenStore = null)
        {
            if (string.IsNullOrEmpty(clientId)) ThrowServiceError(GraphErrorCode.InvalidRequest, ServiceMessages.NoClientId);

            _httpProvider = new HttpProvider();
            _tokenServiceUrl = Endpoints.AuthenticationTokenServiceUrl;
            _redirectUri = Endpoints.NativeRedirectUrl;
            _refreshToken = (tokenStore == null) ? null : tokenStore.RefreshToken;
            _store = tokenStore;
            _clientId = clientId;
            _scopes.Add(GraphScopes.UserRead);
            _appBased = false;
        }

        /// <summary>
        /// Constructor for app based authentication.
        /// </summary>
        /// <param name="clientId">The client id for the application.</param>
        /// <param name="clientSecret">The client secret which will be used to perform application authentication.</param>
        /// <param name="tenantId">The Office 365 tenant id to perform application authentication.</param>
        public AuthenticationProvider(string clientId, string clientSecret, string tenantId)
        {
            if (string.IsNullOrEmpty(clientId)) ThrowServiceError(GraphErrorCode.InvalidRequest, ServiceMessages.NoClientId);
            if (string.IsNullOrEmpty(clientSecret)) ThrowServiceError(GraphErrorCode.InvalidRequest, ServiceMessages.NoClientSecret);
            if (string.IsNullOrEmpty(tenantId)) ThrowServiceError(GraphErrorCode.InvalidRequest, ServiceMessages.NoTenantId);

            _httpProvider = new HttpProvider();
            _tokenServiceUrl = string.Format(Endpoints.TokenServiceUrl, tenantId);
            _clientId = clientId;
            _clientSecret = clientSecret;
            _tenantId = tenantId;
            _scopes.Add(GraphScopes.Default);
            _appBased = true;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Attempts to authenticate either the app or the user.
        /// </summary>
        /// <returns>The task to await.</returns>
        public async Task AuthenticateAsync()
        {
            if (Authenticated) return;

            if (!_appBased && !string.IsNullOrEmpty(_refreshToken))
            { 
                try
                {
                    await RefreshAccessTokenAsync();
                }
                catch
                {
                    if (_store != null) _store.RefreshToken = _refreshToken;

                    _refreshToken = null;
                }

                if (Authenticated) return;
            }

            await GetAuthenticationResultAsync();

            if (!Authenticated) ThrowServiceError(GraphErrorCode.AuthenticationFailure, ServiceMessages.NoToken);
        }

        /// <summary>
        /// Adds the current access token to the request headers. This method will silently refresh the access
        /// token if needed (refresh token is present). If the refresh token is expired or invalid the user
        /// will be prompted to login again.
        /// </summary>
        /// <param name="request">The <see cref="HttpRequestMessage"/> to authenticate.</param>
        /// <returns>The task to await.</returns>
        public async Task AuthenticateRequestAsync(HttpRequestMessage request)
        {
            if (SignRequest(request)) return;

            try
            {
                _accessToken = null;

                if (!_appBased && !string.IsNullOrEmpty(_refreshToken))
                {
                    try
                    {
                        await RefreshAccessTokenAsync();
                    }
                    catch
                    {
                        if (_store != null) _store.RefreshToken = null;

                        _refreshToken = null;
                    }

                    if (!string.IsNullOrEmpty(_accessToken)) return;
                }

                await GetAuthenticationResultAsync();
            }
            finally
            {
                if (!SignRequest(request)) ThrowServiceError(GraphErrorCode.AuthenticationFailure, ServiceMessages.NoToken);
            }
        }

        /// <summary>
        /// Gets the administration consent to allow application to access resources.
        /// </summary>
        /// <param name="redirectUri">The redirect uri for the oAuth request.</param>
        /// <returns>True if consent was given, otherwise false.</returns>
        public bool GetAdminConsent(string redirectUri)
        {
            if (string.IsNullOrEmpty(_tenantId)) ThrowServiceError(GraphErrorCode.InvalidRequest, ServiceMessages.NoTenantId);

            WebAuthenticationBroker.WebAuthenticationBroker.Height = AuthWindowHeight;

            var requestUri = new Uri(GetAdminConsentRequestUrl(redirectUri));
            var result = WebAuthenticationBroker.WebAuthenticationBroker.Authenticate(WebAuthenticationOptions.None, requestUri, new Uri(redirectUri));

            IDictionary<string, string> responseValues = null;

            if ((result != null) && !string.IsNullOrEmpty(result.ResponseData))
            {
                responseValues = UrlHelper.GetQueryOptions(new Uri(result.ResponseData));

                ThrowIfError(responseValues);
            }
            else if ((result != null) && result.ResponseStatus == WebAuthenticationStatus.UserCancel)
            {
                ThrowServiceError(GraphErrorCode.AuthenticationFailure, ServiceMessages.UserCancel);
            }

            var allowed = string.Empty;

            return ((responseValues != null) && responseValues.TryGetValue(Params.AdminConsentKeyName, out allowed)) ? Convert.ToBoolean(allowed) : false;
        }

        /// <summary>
        /// Logs the user or app out of the tenant.
        /// </summary>
        public void Logout()
        {
            _refreshToken = null;

            try
            {
                if (_store != null) _store.RefreshToken = _refreshToken;

                if (Authenticated && !_appBased)
                {
                    WebAuthenticationBroker.WebAuthenticationBroker.Height = AuthWindowHeight;

                    var requestUri = new Uri(GetLogoutRequestUrl(Endpoints.LogoutRedirectUrl));
                    var result = WebAuthenticationBroker.WebAuthenticationBroker.Authenticate(WebAuthenticationOptions.None, requestUri, new Uri(Endpoints.LogoutRedirectUrl));
                }
            }
            finally
            {
                _accessToken = null;
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Returns true if configured for application authentication.
        /// </summary>
        public bool AppBased
        {
            get { return _appBased; }
        }

        /// <summary>
        /// Returns true if the provider is authenticated. 
        /// </summary>
        public bool Authenticated
        {   
            get { return (!string.IsNullOrEmpty(_accessToken) && (_expiration > DateTimeOffset.Now.UtcDateTime.AddMinutes(5))); }
        }

        /// <summary>
        /// Returns the access token.
        /// </summary>
        public string AccessToken
            {
                get { return _accessToken; }
            }

        /// <summary>
        /// Returns the client id.
        /// </summary>
        public string ClientId
        {
            get { return _clientId; }
        }

        /// <summary>
        /// The refresh token.
        /// </summary>
        public string RefreshToken
        {
            get { return _refreshToken; }
            set
            {
                _refreshToken = value;

                if (_store != null) _store.RefreshToken = value;
            }
        }

        /// <summary>
        /// Collection of scopes to request when authenticating.
        /// </summary>
        public IList<string> Scopes
        {
            get { return _scopes; }
        }

        /// <summary>
        /// Returns the tenant id.
        /// </summary>
        public string TenantId
        {
            get { return _tenantId; }
        }

        #endregion
    }
}