/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Sage.Office365.Graph.Authentication;
using Sage.Office365.Graph.Extensions;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Sage.Office365.Graph
{
    /// <summary>
    /// Client class for Sage integration with Office 365 Graph.
    /// </summary>
    public class Client
    {
        #region Private fields

        private readonly static string[] SubFolders = { "Statements", "SalesHistory", "Quotes", "PriceList" };
        private AuthenticationHelper _helper;
        private string _clientId;

        #endregion

        #region Constructor

        /// <summary>
        /// Private constructor.
        /// </summary>
        private Client(string clientId)
        {
            if (string.IsNullOrEmpty(clientId)) throw new ArgumentNullException("clientId");

            _clientId = clientId;
            _helper = new AuthenticationHelper(_clientId);
            _helper.SignIn();
        }

        #endregion

        #region Private methods

        /// <summary>
        /// Executes the task using a method that allows message processing on the calling thread.
        /// </summary>
        /// <typeparam name="T">The data type to return.</typeparam>
        /// <param name="task">The task to execute.</param>
        /// <returns>The result type from the task on success, throws on failure.</returns>
        private T ExecuteTask<T>(Task<T> task)
        {
            if (task == null) throw new ArgumentNullException("task");
            if (!_helper.SignedIn) _helper.SignIn();

            task.WaitWithPumping();

            if (task.IsFaulted)
            {
                if ((task.Exception != null) && (task.Exception.InnerException != null) && (task.Exception.InnerException is ServiceException))
                {
                    var serviceException = (ServiceException)task.Exception.InnerException;

                    if (serviceException.Error != null) throw new ApplicationException(serviceException.Error.ToString());
                }
                throw new ApplicationException(task.Exception.ToString());
            }
            if (task.IsCanceled) throw new OperationCanceledException("The task was cancelled.");

            return task.Result;
        }

        #endregion

        #region Public methods

        /// <summary>
        /// Static factory method for creating the graph client handler.
        /// </summary>
        /// <param name="clientId">The client id to use.</param>
        /// <returns>An instance of client on success, throws on failure.</returns>
        public static Client Create(string clientId)
        {
            return new Client(clientId);
        }

        /// <summary>
        /// Sign the user in to the client.
        /// </summary>
        public void SignIn()
        {
            if (!_helper.SignedIn) _helper.SignIn();
        }

        /// <summary>
        /// Sign the user out of the client.
        /// </summary>
        public void SignOut()
        {
            if (_helper.SignedIn) _helper.SignOut();
        }

        /// <summary>
        /// Get the current user object.
        /// </summary>
        /// <returns>The /Me instance from Office 365.</returns>
        public User Me()
        {
            return ExecuteTask(_helper.Client.Me.Request().GetAsync());
        }

        /// <summary>
        /// Returns a list of all drive items in the root of OneDrive for the logged in user.
        /// </summary>
        /// <returns>The list of drive items.</returns>
        public IList<DriveItem> GetRootItems()
        {
            var rootPage = ExecuteTask(_helper.Client.Me.Drive.Root.Children.Request().GetAsync());
            var result = new List<DriveItem>();

            while (rootPage != null)
            {
                foreach (var item in rootPage.CurrentPage) result.Add(item);

                rootPage = (rootPage.NextPageRequest == null) ? null : ExecuteTask(rootPage.NextPageRequest.GetAsync());
            }

            return result;
        }

        /// <summary>
        /// Attempts to get the root folder item using the drive item name.
        /// </summary>
        /// <param name="rootName">The name of the root item to obtain.</param>
        /// <returns>The drive item on success, throws on failure.</returns>
        public DriveItem GetRootItem(string rootName)
        {
            return ExecuteTask(_helper.Client.Me.Drive.Root.ItemWithPath(rootName).Request().GetAsync());
        }

        /// <summary>
        /// Gets the drive item children for the specified folder drive item.
        /// </summary>
        /// <param name="folder">The folder to obtain the children for.</param>
        /// <returns>The list of drive item children on success, throws on failure.</returns>
        public IList<DriveItem> GetChildren(DriveItem folder)
        {
            if (folder == null) throw new ArgumentNullException("folder");

            var childPage = ExecuteTask(_helper.Client.Me.Drive.Items[folder.Id].Children.Request().GetAsync());
            var result = new List<DriveItem>();

            while (childPage != null)
            {
                foreach (var item in childPage.CurrentPage) result.Add(item);

                childPage = (childPage.NextPageRequest == null) ? null : ExecuteTask(childPage.NextPageRequest.GetAsync());
            }

            return result;
        }

        #endregion

        #region Public properties

        /// <summary>
        /// The client id that the client was initialized with.
        /// </summary>
        public string ClientId
        {
            get {  return _clientId; }
        }

        /// <summary>
        /// Returns true if the user is signed into Office 365.
        /// </summary>
        public bool SignedIn
        {
            get { return _helper.SignedIn; }
        }

        #endregion
    }
}
