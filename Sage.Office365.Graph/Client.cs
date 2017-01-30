/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Sage.Office365.Graph.Authentication;
using Sage.Office365.Graph.Authentication.Interfaces;
using Sage.Office365.Graph.Authentication.Storage;
using Sage.Office365.Graph.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace Sage.Office365.Graph
{
    /// <summary>
    /// Client class for Sage integration with Office 365 OneDrive
    /// </summary>
    public class Client
    {
        #region Private fields

        private IAuthenticationProvider2 _provider;
        private GraphServiceClient _serviceClient;
        private User _principal;

        #endregion

        #region Private methods

        /// <summary> 
        /// Reads a file fragment from the stream.
        /// </summary> 
        /// <param name="stream">The file stream to read from.</param> 
        /// <param name="position">The position to start reading from.</param> 
        /// <param name="count">The number of bytes to read.</param> 
        /// <returns>The fragment of file with byte[].</returns> 
        private byte[] ReadFileFragment(Stream stream, long position, int count)
        {
            if ((position >= stream.Length) || (position < 0) || (count <= 0)) return null;

            var trimCount = ((position + count) > stream.Length) ? (stream.Length - position) : count;
            var result = new byte[trimCount];

            stream.Seek(position, SeekOrigin.Begin);
            stream.Read(result, 0, (int)trimCount);

            return result;
        }

        /// <summary>
        /// Formats the drive path by removing any leading "/drive/root:" reference, as well as any
        /// leading '/' character.
        /// </summary>
        /// <param name="pathName">The path name to format.</param>
        /// <returns>The formated path name.</returns>
        private string FormatDrivePath(string pathName)
        {
            if (string.IsNullOrEmpty(pathName)) throw new ArgumentNullException("pathName");

            var temp = (pathName.StartsWith(Common.Constants.DriveRoot, StringComparison.OrdinalIgnoreCase)) ? pathName.Remove(0, Common.Constants.DriveRoot.Length) : pathName;

            return (temp.StartsWith("/")) ? temp.Remove(0, 1) : temp;
        }

        /// <summary>
        /// Returns a list of all drive items in the root of One Drive for the logged in user.
        /// </summary>
        /// <returns>The list of drive items.</returns>
        private IList<DriveItem> GetRootItems()
        {
            SignIn();

            var rootPage = ExecuteTask(_serviceClient.Me.Drive.Root.Children.Request().GetAsync());
            var result = new List<DriveItem>();

            while (rootPage != null)
            {
                foreach (var item in rootPage.CurrentPage) result.Add(item);

                rootPage = (rootPage.NextPageRequest == null) ? null : ExecuteTask(rootPage.NextPageRequest.GetAsync());
            }

            return result;
        }

        /// <summary>
        /// Special handling for Graph service exceptions, where the Error field is more informative
        /// than the exception itself.
        /// </summary>
        /// <param name="exception">The exception to process.</param>
        private void ThrowException(Exception exception)
        {
            if (exception == null) return;
            if ((exception.InnerException != null) && (exception.InnerException is ServiceException))
            {
                var serviceException = (ServiceException)exception.InnerException;

                if (serviceException.Error != null) throw new ApplicationException(serviceException.Error.ToString());
            }

            throw new ApplicationException(exception.ToString());
        }

        /// <summary>
        /// Executes the task using a method that allows message processing on the calling thread.
        /// </summary>
        /// <param name="task">The task to execute.</param>
        private void ExecuteTask(Task task)
        {
            if (task == null) throw new ArgumentNullException("task");

            task.WaitWithPumping();

            if (task.IsFaulted) ThrowException(task.Exception);
            if (task.IsCanceled) throw new OperationCanceledException("The task was cancelled.");
        }

        /// <summary>
        /// Executes the task using a method that allows message processing on the calling thread.
        /// </summary>
        /// <typeparam name="T">The data type to return.</typeparam>
        /// <param name="task">The task to execute.</param>
        /// <returns>The result type from the task on success, throws on failure.</returns>
        private T ExecuteTask<T>(Task<T> task)
        {
            if (task == null) throw new ArgumentNullException("task");

            task.WaitWithPumping();

            if (task.IsFaulted) ThrowException(task.Exception);
            if (task.IsCanceled) throw new OperationCanceledException("The task was cancelled.");

            return task.Result;
        }

        #endregion

        #region Constructor

        /// <summary>
        /// Constructor for user based authentication.
        /// </summary>
        /// <param name="clientId">The client id for the application.</param>
        public Client(string clientId)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");

            _provider = new AuthenticationProvider(clientId, new DesktopAuthenticationStore(clientId, Scope.System));
        }

        /// <summary>
        /// Constructor for app based authentication.
        /// </summary>
        /// <param name="clientId">The client id for the application.</param>
        /// <param name="clientSecret">The client secret which will be used to perform application authentication.</param>
        /// <param name="tenantId">The Office 365 tenant id to perform application authentication.</param>
        public Client(string clientId, string clientSecret, string tenantId)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");
            if (string.IsNullOrEmpty(clientSecret)) throw new ArgumentNullException("clientSecret");
            if (string.IsNullOrEmpty(tenantId)) throw new ArgumentNullException("tenantId");

            _provider = new AuthenticationProvider(clientId, clientSecret, tenantId);
       }

        #endregion

        #region Public methods

        /// <summary>
        /// Sign the user in to the client.
        /// </summary>
        public void SignIn()
        {
            if (SignedIn) return;

            try
            {
                if (!_provider.Authenticated) ExecuteTask(_provider.AuthenticateAsync());

                _serviceClient = new GraphServiceClient(_provider);

                if (!_provider.AppBased) _principal = ExecuteTask(_serviceClient.Me.Request().GetAsync());
            }
            catch (Exception)
            {
                _serviceClient = null;

                throw;
            }
        }

        /// <summary>
        /// Sign the user out of the client.
        /// </summary>
        public void SignOut()
        {
            try
            {
                if (_provider.Authenticated) _provider.Logout();
            }
            finally
            {
                _principal = null;
                _serviceClient = null;
            }
        }

        /// <summary>
        /// Get the collection of users.
        /// </summary>
        /// <returns>The collection of users.</returns>
        public IList<User> Users()
        {
            SignIn();

            var result = new List<User>();
            var userPage = ExecuteTask(_serviceClient.Users.Request().GetAsync());

            while (userPage != null)
            {
                foreach (var user in userPage.CurrentPage)
                {
                    result.Add(user);
                }

                userPage = (userPage.NextPageRequest == null) ? null : ExecuteTask(userPage.NextPageRequest.GetAsync());
            }

            return result;
        }

        #endregion

        #region Public properties

        /// <summary>
        /// Returns the authentication provider.
        /// </summary>
        public IAuthenticationProvider2 Provider
        {
            get { return _provider; }
        }

        /// <summary>
        /// Returns the instance of the graph service client.
        /// </summary>
        public GraphServiceClient ServiceClient
        {
            get { return _serviceClient; }
        }

        /// <summary>
        /// The current user princial.
        /// </summary>
        public User Principal
        {
            get { return _principal; }
            set { _principal = value; }
        }

        /// <summary>
        /// Returns true if the user is signed into the graph service client.
        /// </summary>
        public bool SignedIn
        {
            get { return ((_serviceClient != null) && _provider.Authenticated); }
        }

        #endregion
    }
}
