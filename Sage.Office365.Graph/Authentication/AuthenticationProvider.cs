/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
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
    /// Client based authentication provider for Microsoft Graph.
    /// </summary>
    [SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable")]
    public class OAuth2AuthenticationProvider : IAuthenticationProvider
    {
        #region Private fields

        private IHttpProvider _httpProvider;
        private DateTimeOffset _expiration;
        private IAuthenticationStore _store;
        private string _clientId;
        private string[] _scopes;
        private string _returnUrl;
        private string _accessToken;
        private string _refreshToken;

        #endregion

        #region Private methods

        /// <summary>
        /// Gets the authentication result.
        /// </summary>
        /// <returns>The task to wait on.</returns>
        private async Task GetAuthenticationResultAsync()
        {
            var code = GetAuthorizationCodeAsync(_returnUrl);

            if (!string.IsNullOrEmpty(code))
            {
                await SendTokenRequestAsync(GetCodeRedemptionRequestBody(code));
            }
        }

        /// <summary>
        /// Gets the authorization code.
        /// </summary>
        /// <param name="returnUrl">The redirect uri for the oAuth request.</param>
        /// <returns>The oAuth code.</returns>
        private string GetAuthorizationCodeAsync(string returnUrl = null)
        {
            var requestUri = new Uri(GetAuthorizationCodeRequestUrl(returnUrl));
            var result = WebAuthenticationBroker.WebAuthenticationBroker.Authenticate(WebAuthenticationOptions.None, requestUri, new Uri(_returnUrl));

            IDictionary<string, string> authenticationResponseValues = null;

            if ((result != null) && !string.IsNullOrEmpty(result.ResponseData))
            {
                authenticationResponseValues = UrlHelper.GetQueryOptions(new Uri(result.ResponseData));

                ThrowIfError(authenticationResponseValues);
            }
            else if ((result != null) && result.ResponseStatus == WebAuthenticationStatus.UserCancel)
            {
                return null;
            }

            var code = string.Empty;

            return ((authenticationResponseValues != null) && authenticationResponseValues.TryGetValue("code", out code)) ? code : null;
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
        /// Sends the token request.
        /// </summary>
        /// <param name="requestBodyString">The request body.</param>
        /// <returns>The task to wait on.</returns>
        private async Task SendTokenRequestAsync(string requestBodyString)
        {
            var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, OAuth2AuthenticationProvider.TokenServiceUrl);

            httpRequestMessage.Content = new StringContent(requestBodyString, Encoding.UTF8, Headers.FormUrlEncodedContentType);

            using (var authResponse = await _httpProvider.SendAsync(httpRequestMessage))
            {
                using (var responseStream = await authResponse.Content.ReadAsStreamAsync())
                {
                    var responseValues = _httpProvider.Serializer.DeserializeObject<IDictionary<string, string>>(responseStream);

                    if (responseValues != null)
                    {
                        ThrowIfError(responseValues);

                        _refreshToken = responseValues[Authentication.RefreshTokenKeyName];
                        _accessToken = responseValues[Authentication.AccessTokenKeyName];
                        _expiration = DateTimeOffset.UtcNow.Add(new TimeSpan(0, 0, int.Parse(responseValues[Authentication.ExpiresInKeyName])));
                        _store.RefreshToken = _refreshToken;
                            
                        return;
                    }

                    throw new ServiceException(
                        new Error
                        {
                            Code = GraphErrorCode.AuthenticationFailure.ToString(),
                            Message = "Authentication failed. No response values returned from token authentication flow."
                        });
                }
            }
        }

        /// <summary>
        /// Gets the request URL for OAuth authentication using the code flow.
        /// </summary>
        /// <param name="returnUrl">The return URL for the request. Defaults to the service info value.</param>
        /// <returns>The OAuth request URL.</returns>
        private string GetAuthorizationCodeRequestUrl(string returnUrl = null)
        {
            var requestUriStringBuilder = new StringBuilder();

            requestUriStringBuilder.Append(AuthenticationServiceUrl);
            requestUriStringBuilder.AppendFormat("?{0}={1}", Authentication.RedirectUriKeyName, returnUrl);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Authentication.ClientIdKeyName, _clientId);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Authentication.ResponseTypeKeyName, Authentication.CodeKeyName);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Authentication.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", _scopes)));

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

            requestBodyStringBuilder.AppendFormat("{0}={1}", Authentication.RedirectUriKeyName, _returnUrl);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.ClientIdKeyName, _clientId);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.CodeKeyName, code);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.GrantTypeKeyName, Authentication.AuthorizationCodeGrantType);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", _scopes)));

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

            requestBodyStringBuilder.AppendFormat("{0}={1}", Authentication.RedirectUriKeyName, _returnUrl);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.ClientIdKeyName, _clientId);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.RefreshTokenKeyName, refreshToken);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.GrantTypeKeyName, Authentication.RefreshTokenKeyName);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", _scopes)));

            return requestBodyStringBuilder.ToString();
        }

        /// <summary>
        /// Throws an exception if an error is found in the dictionary.
        /// </summary>
        /// <param name="responseValues">The oAuth response values.</param>
        private void ThrowIfError(IDictionary<string, string> responseValues)
        {
            if (responseValues == null) return;

            string error = null;
            string errorDescription = null;

            if (responseValues.TryGetValue(Authentication.ErrorDescriptionKeyName, out errorDescription) || responseValues.TryGetValue(Authentication.ErrorKeyName, out error))
            {
                ParseAuthenticationError(error, errorDescription);
            }
        }

        /// <summary>
        /// Parses the authentication error and throws an exception.
        /// </summary>
        /// <param name="error">The error.</param>
        /// <param name="errorDescription">The error description.</param>
        private void ParseAuthenticationError(string error, string errorDescription)
        {
            throw new ServiceException(
                new Error
                {
                    Code = GraphErrorCode.AuthenticationFailure.ToString(),
                    Message = errorDescription ?? error
                });
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

            request.Headers.Authorization = new AuthenticationHeaderValue("bearer", _accessToken);

            return true;
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="store">The DI refresh token store to use.</param>
        /// <param name="clientId">The client id for the application.</param>
        /// <param name="returnUrl">The redirect uri for authentication.</param>
        /// <param name="scopes">The requested scopes.</param>
        /// <param name="httpProvider">The optional HttpProvider.</param>
        public OAuth2AuthenticationProvider(IAuthenticationStore store, string clientId, string returnUrl, string[] scopes, IHttpProvider httpProvider = null)
        {
            if (store == null)
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.ServiceNotAvailable.ToString(),
                        Message = "An authentication store is required when using OAuth2AuthenticationProvider.",
                    });
            }

            if (string.IsNullOrEmpty(clientId))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "Client ID is required to authenticate using OAuth2AuthenticationProvider.",
                    });
            }

            if (string.IsNullOrEmpty(returnUrl))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "Return URL is required to authenticate using OAuth2AuthenticationProvider.",
                    });
            }

            if (scopes == null || scopes.Length == 0)
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "No scopes have been requested for authentication.",
                    });
            }

            _store = store;
            _refreshToken = _store.RefreshToken;
            _httpProvider = httpProvider ?? new HttpProvider();
            _clientId = clientId;
            _returnUrl = returnUrl;
            _scopes = scopes;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Presents authentication UI to the user.
        /// </summary>
        /// <returns>The task to await.</returns>
        public async Task AuthenticateAsync()
        {
            if (!string.IsNullOrEmpty(_accessToken) && !(_expiration <= DateTimeOffset.Now.UtcDateTime.AddMinutes(5))) return;
            if (!string.IsNullOrEmpty(_refreshToken))
            {
                try
                {
                    await RefreshAccessTokenAsync();
                }
                catch
                {
                    _refreshToken = null;
                    _store.RefreshToken = _refreshToken;
                }
            }

            if (!string.IsNullOrEmpty(_accessToken) && !(_expiration <= DateTimeOffset.Now.UtcDateTime.AddMinutes(5))) return;

            await GetAuthenticationResultAsync();

            if (string.IsNullOrEmpty(_accessToken))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.AuthenticationFailure.ToString(),
                        Message = "Failed to retrieve a valid authentication token for the user."
                    });
            }
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

            _accessToken = null;

            try
            {
                try
                {
                    await RefreshAccessTokenAsync();
                }
                catch
                {
                    _store.RefreshToken = _refreshToken = null;
                }

                if (!string.IsNullOrEmpty(_accessToken)) return;

                await GetAuthenticationResultAsync();
            }
            finally
            {
                if (!SignRequest(request))
                { 
                    throw new ServiceException(
                        new Error
                        {
                            Code = GraphErrorCode.AuthenticationFailure.ToString(),
                            Message = "Failed to retrieve a valid authentication token for the user."
                        });
                }
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Returns the access token.
        /// </summary>
        public string AccessToken
        {
            get { return _accessToken; }
        }

        /// <summary>
        /// The refresh token.
        /// </summary>
        public string RefreshToken
        {
            get { return _refreshToken; }
            set
            {
                _store.RefreshToken = _refreshToken = value;
            }
        }

        #endregion
    }

    /// <summary>
    /// Application authentication provider for Microsoft Graph.
    /// </summary>
    [SuppressMessage("Microsoft.Design", "CA1001:TypesThatOwnDisposableFieldsShouldBeDisposable")]
    public class OAuth2AppAuthenticationProvider : IAuthenticationProvider
    {
        #region Private fields

        private IHttpProvider _httpProvider;
        private DateTimeOffset _expiration;
        private readonly static string[] _scopes = { "https://graph.microsoft.com/.default" };
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _tenantId;
        private readonly string _redirectUri;
        private string _accessToken;

        #endregion

        #region Private methods

        /// <summary>
        /// Sends the token request.
        /// </summary>
        /// <param name="requestBodyString">The request body.</param>
        /// <returns>The task to wait on.</returns>
        private async Task SendTokenRequestAsync(string requestBodyString)
        {
            var tokenServiceUrl = string.Format(OAuth2AppAuthenticationProvider.TokenServiceUrl, _tenantId);
            var httpRequestMessage = new HttpRequestMessage(HttpMethod.Post, tokenServiceUrl);

            httpRequestMessage.Content = new StringContent(requestBodyString, Encoding.UTF8, Headers.FormUrlEncodedContentType);

            using (var authResponse = await _httpProvider.SendAsync(httpRequestMessage))
            {
                using (var responseStream = await authResponse.Content.ReadAsStreamAsync())
                {
                    var responseValues = _httpProvider.Serializer.DeserializeObject<IDictionary<string, string>>(responseStream);

                    if (responseValues != null)
                    {
                        ThrowIfError(responseValues);

                        _accessToken = responseValues[Authentication.AccessTokenKeyName];
                        _expiration = DateTimeOffset.UtcNow.Add(new TimeSpan(0, 0, int.Parse(responseValues[Authentication.ExpiresInKeyName])));

                        return;
                    }

                    throw new ServiceException(
                        new Error
                        {
                            Code = GraphErrorCode.AuthenticationFailure.ToString(),
                            Message = "Authentication failed. No response values returned from token authentication flow."
                        });
                }
            }
        }

        /// <summary>
        /// Gets the administration consent url.
        /// </summary>
        /// <param name="returnUrl">The return URL for the request. Defaults to the service info value.</param>
        /// <returns>The OAuth request URL.</returns>
        private string GetAdminConsentRequestUrl(string returnUrl = null)
        {
            var requestUriStringBuilder = new StringBuilder();

            requestUriStringBuilder.Append(string.Format(AdminConsentUrl, _tenantId));
            requestUriStringBuilder.AppendFormat("?{0}={1}", Authentication.ClientIdKeyName, _clientId);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Authentication.StateKeyName, ConsentState);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Authentication.RedirectUriKeyName, returnUrl);

            return requestUriStringBuilder.ToString();
        }

        private string GetLogoutRequestUrl(string returnUrl = null)
        {
            var requestUriStringBuilder = new StringBuilder();

            requestUriStringBuilder.Append(LogoutServiceUrl);
            requestUriStringBuilder.AppendFormat("?{0}={1}", Authentication.ClientIdKeyName, _clientId);
            requestUriStringBuilder.AppendFormat("&{0}={1}", Authentication.PostRedirectUriKeyName, returnUrl);

            return requestUriStringBuilder.ToString();
        }

        private bool LogoutUser(string returnUrl)
        {
            var requestUri = new Uri(GetLogoutRequestUrl(returnUrl));
            var result = WebAuthenticationBroker.WebAuthenticationBroker.Authenticate(WebAuthenticationOptions.None, requestUri, new Uri(returnUrl));

            IDictionary<string, string> responseValues = null;

            if ((result != null) && !string.IsNullOrEmpty(result.ResponseData))
            {
                responseValues = UrlHelper.GetQueryOptions(new Uri(result.ResponseData));

                ThrowIfError(responseValues);
            }
            else if ((result != null) && result.ResponseStatus == WebAuthenticationStatus.UserCancel)
            {
                return false;
            }

            var allowed = string.Empty;

            return ((responseValues != null) && responseValues.TryGetValue("admin_consent", out allowed)) ? Convert.ToBoolean(allowed) : false;
        }

        /// <summary>
        /// Gets the administration consent to allow application to access resources.
        /// </summary>
        /// <param name="returnUrl">The redirect uri for the oAuth request.</param>
        /// <returns>True if consent was given, otherwise false.</returns>
        private bool GetAdminConsent(string returnUrl = null)
        {
            var requestUri = new Uri(GetAdminConsentRequestUrl(returnUrl));
            var result = WebAuthenticationBroker.WebAuthenticationBroker.Authenticate(WebAuthenticationOptions.None, requestUri, new Uri(returnUrl));

            IDictionary<string, string> responseValues = null;

            if ((result != null) && !string.IsNullOrEmpty(result.ResponseData))
            {
                responseValues = UrlHelper.GetQueryOptions(new Uri(result.ResponseData));

                ThrowIfError(responseValues);
            }
            else if ((result != null) && result.ResponseStatus == WebAuthenticationStatus.UserCancel)
            {
                return false;
            }

            var allowed = string.Empty;

            return ((responseValues != null) && responseValues.TryGetValue("admin_consent", out allowed)) ? Convert.ToBoolean(allowed) : false;
        }

        /// <summary>
        /// Gets the request body for generating an application based access token.
        /// </summary>
        /// <returns>The request body for the call.</returns>
        private string GetAccessTokenRequestBody()
        {
            var requestBodyStringBuilder = new StringBuilder();

            requestBodyStringBuilder.AppendFormat("{0}={1}", Authentication.ClientIdKeyName, _clientId);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.ClientSecretKeyName, _clientSecret);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.GrantTypeKeyName, Authentication.ClientCredentions);
            requestBodyStringBuilder.AppendFormat("&{0}={1}", Authentication.ScopeKeyName, WebUtility.UrlEncode(string.Join(" ", _scopes)));

            return requestBodyStringBuilder.ToString();
        }

        /// <summary>
        /// Throws an exception if an error is found in the dictionary.
        /// </summary>
        /// <param name="responseValues">The oAuth response values.</param>
        private void ThrowIfError(IDictionary<string, string> responseValues)
        {
            if (responseValues == null) return;

            string error = null;
            string errorDescription = null;

            if (responseValues.TryGetValue(Authentication.ErrorDescriptionKeyName, out errorDescription) || responseValues.TryGetValue(Authentication.ErrorKeyName, out error))
            {
                ParseAuthenticationError(error, errorDescription);
            }
        }

        /// <summary>
        /// Parses the authentication error and throws an exception.
        /// </summary>
        /// <param name="error">The error.</param>
        /// <param name="errorDescription">The error description.</param>
        private void ParseAuthenticationError(string error, string errorDescription)
        {
            throw new ServiceException(
                new Error
                {
                    Code = GraphErrorCode.AuthenticationFailure.ToString(),
                    Message = errorDescription ?? error
                });
        }

        /// <summary>
        /// Gets the authentication result.
        /// </summary>
        /// <returns>The task to wait on.</returns>
        private async Task GetAuthenticationResultAsync()
        {
            await SendTokenRequestAsync(GetAccessTokenRequestBody());
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

            request.Headers.Authorization = new AuthenticationHeaderValue("bearer", _accessToken);

            return true;
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="clientId">The client id for the application.</param>
        /// <param name="clientSecret">The client secret which will be used to perform application authentication.</param>
        /// <param name="tenantId">The Office 365 tenant id to perform application authentication.</param>
        /// <param name="httpProvider">The optional HttpProvider.</param>
        public OAuth2AppAuthenticationProvider(string clientId, string clientSecret, string tenantId, IHttpProvider httpProvider = null)
        {
            if (string.IsNullOrEmpty(clientId))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "Client ID is required to authenticate using OAuth2AppAuthenticationProvider.",
                    });
            }

            if (string.IsNullOrEmpty(clientSecret))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "Client secret is required to authenticate using OAuth2AppAuthenticationProvider.",
                    });
            }


            if (string.IsNullOrEmpty(tenantId))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.InvalidRequest.ToString(),
                        Message = "Tenant ID is required to authenticate using OAuth2AppAuthenticationProvider.",
                    });
            }

            _httpProvider = httpProvider ?? new HttpProvider();
            _redirectUri = "http://localhost/sagepaperless";
            _clientId = clientId;
            _clientSecret = clientSecret;
            _tenantId = tenantId;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Silently attempts to authenticate the user.
        /// </summary>
        /// <returns>The task to await.</returns>
        public async Task AuthenticateAsync()
        {
            if (!string.IsNullOrEmpty(_accessToken) && !(_expiration <= DateTimeOffset.Now.UtcDateTime.AddMinutes(5))) return;

            var signedOut = LogoutUser("http://localhost/sagepaperless");

            var consent = GetAdminConsent(_redirectUri);
            await GetAuthenticationResultAsync();

            if (string.IsNullOrEmpty(_accessToken))
            {
                throw new ServiceException(
                    new Error
                    {
                        Code = GraphErrorCode.AuthenticationFailure.ToString(),
                        Message = "Failed to retrieve a valid authentication token for the user."
                    });
            }
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

            _accessToken = null;

            try
            {
                await GetAuthenticationResultAsync();
            }
            finally
            {
                if (!SignRequest(request))
                {
                    throw new ServiceException(
                        new Error
                        {
                            Code = GraphErrorCode.AuthenticationFailure.ToString(),
                            Message = "Failed to retrieve a valid authentication token for the user."
                        });
                }
            }
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Returns the access token.
        /// </summary>
        public string AccessToken
        {
            get { return _accessToken; }
        }

        #endregion
    }
}
