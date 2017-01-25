/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Sage.Office365.Graph.Authentication;
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
    /// Client class for Sage integration with Office 365 Graph.
    /// </summary>
    public class Client
    {
        #region Private fields

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
            _helper = new AuthenticationHelper(new DesktopAuthenticationStore(_clientId, Scope.System), _clientId);

            SignIn();
        }

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
            if (_helper.SignedIn) return;

            try
            {
                _helper.SignIn();
            }
            catch (Exception exception)
            {
                ThrowException(exception);
            }
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
            SignIn();

            return ExecuteTask(_helper.Client.Me.Request().GetAsync());
        }

        /// <summary>
        /// Attempts to get the drive item using the specified path, which may contain any number of sub folders
        /// seperated by the '/' character.
        /// </summary>
        /// <param name="pathName">The qualified path of the drive item to obtain.</param>
        /// <returns>The drive item on success, throws on failure.</returns>
        /// <example>
        /// var pdf = Client.GetItem("3031414246/SalesHistory/Accounts Receivable Invoice History Report.pdf");
        /// </example>
        public DriveItem GetItem(string pathName)
        {
            var qualifiedName = FormatDrivePath(pathName);

            if (string.IsNullOrEmpty(qualifiedName)) throw new ApplicationException("Formatted path name cannot be null.");

            SignIn();

            var c = ExecuteTask(_helper.Client.Me.Drive.Root.ItemWithPath(qualifiedName).Request().GetAsync());

            return c;
        }

        /// <summary>
        /// Returns the parent folder for the specified drive item.
        /// </summary>
        /// <param name="folderOrFile">The item to return the parent for.</param>
        /// <returns>The parent drive item on success, throws on failure.</returns>
        public DriveItem GetParent(DriveItem folderOrFile)
        {
            if (folderOrFile == null) throw new ArgumentNullException("item");
            if (folderOrFile.ParentReference == null) return null;

            var temp = FormatDrivePath(folderOrFile.ParentReference.Path);

            if (string.IsNullOrEmpty(temp)) return ExecuteTask(_helper.Client.Me.Drive.Root.Request().GetAsync());

            return GetItem(temp);
        }

        /// <summary>
        /// Gets the drive item children for the specified folder drive item. If folder is null,
        /// then the children for the root of the drive are returned.
        /// </summary>
        /// <param name="folder">The folder to obtain the children for.</param>
        /// <returns>The list of drive item children on success, throws on failure.</returns>
        public IList<DriveItem> GetChildren(DriveItem folder = null)
        {
            if (folder == null) return GetRootItems();

            SignIn();

            var childPage = ExecuteTask(_helper.Client.Me.Drive.Items[folder.Id].Children.Request().GetAsync());
            var result = new List<DriveItem>();

            while (childPage != null)
            {
                foreach (var item in childPage.CurrentPage) result.Add(item);

                childPage = (childPage.NextPageRequest == null) ? null : ExecuteTask(childPage.NextPageRequest.GetAsync());
            }

            return result;
        }

        /// <summary>
        /// Returns the qualified path for the drive item.
        /// </summary>
        /// <param name="folderOrFile">The drive item to get the qualified path for.</param>
        /// <returns>The drive item's qualified path.</returns>
        public string GetQualifiedPath(DriveItem folderOrFile)
        {
            if (folderOrFile == null) throw new ArgumentNullException("item");
            if (folderOrFile.ParentReference == null) return folderOrFile.Name;

            var parent = folderOrFile.ParentReference;
            var temp = FormatDrivePath(parent.Path);

            return (string.IsNullOrEmpty(temp) ? folderOrFile.Name : string.Format("{0}/{1}", temp, folderOrFile.Name));
        }

        /// <summary>
        /// Downloads the file from One Drive into the folder specified by localPath. 
        /// </summary>
        /// <param name="file">The One Drive file to download.</param>
        /// <param name="localPath">The local path to download the file to.</param>
        /// <param name="overWriteIfExists">Overwrites an existing file if set to true.</param>
        public void DownloadFile(DriveItem file, string localPath, bool overWriteIfExists = true)
        {
            if ((file == null) || (file.File == null)) throw new ArgumentNullException("file");
            if (string.IsNullOrEmpty(localPath)) throw new ArgumentNullException("localPath");
            if (!Directory.Exists(localPath)) Directory.CreateDirectory(localPath);

            var localFile = Path.Combine(localPath, file.Name);

            if (System.IO.File.Exists(localFile) && !overWriteIfExists) throw new IOException(string.Format("The file '{0}' already exists and overWriteIfExists was set to false.", localFile));

            using (var stream = new FileStream(localFile, FileMode.Create))
            {
                using (var content = ExecuteTask(_helper.Client.Me.Drive.Items[file.Id].Content.Request().GetAsync()))
                {
                    content.CopyTo(stream);
                }
            }
        }

        /// <summary>
        /// Uploads a local file into the specified One Drive folder.
        /// </summary>
        /// <param name="folder">The One Drive folder to upload into.</param>
        /// <param name="localFile">The local file to upload.</param>
        /// <returns>The drive item for the newly uploaded file.</returns>
        public DriveItem UploadFile(DriveItem folder, string localFile)
        {
            if ((folder == null) || (folder.Folder == null)) throw new ArgumentNullException("folder");
            if (string.IsNullOrEmpty(localFile)) throw new ArgumentNullException("localFile");
            if (!System.IO.File.Exists(localFile)) throw new FileNotFoundException();

            using (var stream = new FileStream(localFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var position = 0L;
                var totalLength = stream.Length;
                var remotePath = GetQualifiedPath(folder);
                var remoteFile = (string.IsNullOrEmpty(remotePath) ? Path.GetFileName(localFile) : string.Format("{0}/{1}", remotePath, Path.GetFileName(localFile)));
                var session = ExecuteTask(_helper.Client.Me.Drive.Root.ItemWithPath(remoteFile).CreateUploadSession().Request().PostAsync());

                while (position < totalLength)
                {
                    var bytes = ReadFileFragment(stream, position, Common.Constants.MaxChunkSize);

                    using (var request = new HttpRequestMessage(HttpMethod.Put, session.UploadUrl))
                    {
                        ExecuteTask(_helper.Client.AuthenticationProvider.AuthenticateRequestAsync(request));

                        request.Content = new ByteArrayContent(bytes);
                        request.Content.Headers.ContentRange = new ContentRangeHeaderValue(position, (position + bytes.Length) - 1, totalLength);

                        position += bytes.Length;

                        using (var client = new HttpClient())
                        {
                            using (var response = ExecuteTask(client.SendAsync(request)))
                            {
                                response.EnsureSuccessStatusCode();
                            }
                        }
                    }
                }

                return GetItem(remoteFile);
            }
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
