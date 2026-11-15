using System;
using System.Collections.Concurrent;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Runs all COM work on one dedicated STA thread, strictly one item after another.
    /// Office automation and callbacks from VBA require an STA thread, and AccUnit must never
    /// run tests or VBA calls in parallel against one Access instance.
    /// </summary>
    public sealed class StaComDispatcher : IDisposable
    {
        private const string ThreadName = "AccUnit.McpServer.STA";

        private readonly BlockingCollection<Action> _workItems = new BlockingCollection<Action>();
        private readonly Thread _thread;
        private bool _disposed;

        public StaComDispatcher()
        {
            _thread = new Thread(ProcessWorkItems) { Name = ThreadName, IsBackground = true };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        public bool IsDispatcherThread
        {
            get { return Thread.CurrentThread == _thread; }
        }

        public T Invoke<T>(Func<T> work)
        {
            if (work == null) throw new ArgumentNullException(nameof(work));
            if (_disposed) throw new ObjectDisposedException(nameof(StaComDispatcher));

            // Re-entrant calls would dead-lock the queue; run them directly.
            if (IsDispatcherThread)
                return work();

            var result = default(T);
            ExceptionDispatchInfo error = null;
            using (var done = new ManualResetEventSlim(false))
            {
                _workItems.Add(() =>
                {
                    try { result = work(); }
                    catch (Exception ex) { error = ExceptionDispatchInfo.Capture(ex); }
                    finally { done.Set(); }
                });
                done.Wait();
            }

            if (error != null)
                error.Throw();

            return result;
        }

        public void Invoke(Action work)
        {
            if (work == null) throw new ArgumentNullException(nameof(work));
            Invoke<object>(() => { work(); return null; });
        }

        private void ProcessWorkItems()
        {
            // Managed waits on an STA thread pump COM messages, so callbacks from VBA are delivered.
            foreach (var workItem in _workItems.GetConsumingEnumerable())
            {
                workItem();
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _workItems.CompleteAdding();
            if (!IsDispatcherThread)
                _thread.Join();
            _workItems.Dispose();
        }
    }
}
