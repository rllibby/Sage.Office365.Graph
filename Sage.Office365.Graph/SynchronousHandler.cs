/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using Microsoft.Graph;
using Sage.Office365.Graph.Extensions;
using System;
using System.Threading.Tasks;

namespace Sage.Office365.Graph
{
    /// <summary>
    /// Base level class for executing tasks in a synchronous manner.
    /// </summary>
    public class SynchronousHandler
    {
        #region Public methods

        /// <summary>
        /// Special handling for Graph service exceptions, where the Error field is more informative
        /// than the exception itself.
        /// </summary>
        /// <param name="exception">The exception to process.</param>
        public virtual void ThrowException(Exception exception)
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
        public virtual void ExecuteTask(Task task)
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
        public virtual T ExecuteTask<T>(Task<T> task)
        {
            if (task == null) throw new ArgumentNullException("task");

            task.WaitWithPumping();

            if (task.IsFaulted) ThrowException(task.Exception);
            if (task.IsCanceled) throw new OperationCanceledException("The task was cancelled.");

            return task.Result;
        }

        #endregion
    }
}
