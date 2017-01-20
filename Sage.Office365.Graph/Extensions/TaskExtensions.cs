/*
 *  Copyright © 2017, Sage Software, Inc.
 *  Authored by rllibby.
 */

using System;
using System.Threading.Tasks;
using System.Windows.Threading;

namespace Sage.Office365.Graph.Extensions
{
    /// <summary>
    /// Extension class for Tasks.
    /// </summary>
    public static class TaskExtensions
    {
        /// <summary>
        /// Waits on the task and provides message processing. 
        /// </summary>
        /// <param name="task">The task to wait on.</param>
        public static void WaitWithPumping(this Task task)
        {
            if (task == null) throw new ArgumentNullException("task");

            var nestedFrame = new DispatcherFrame();

            task.ContinueWith(t => nestedFrame.Continue = false);

            Dispatcher.PushFrame(nestedFrame);
        }
    }
}
