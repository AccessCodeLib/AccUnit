using System;

namespace AccessCodeLib.Common.Tools.Logging
{
    public class BlockLogger : IDisposable
    {
        public BlockLogger()
        {
            Info = null;
            LogBlockStart();
        }

        public BlockLogger(string info)
        {
            Info = info;
            LogBlockStart();
        }

        private string Info { get; }

        private void LogBlockStart()
        {
            var message = "Block entry";
            if (!string.IsNullOrEmpty(Info))
            {
                message += ": " + Info;
            }
            Logger.Log(message, 2);
            Logger.Indent();
        }

        private static void LogBlockEnd()
        {
            Logger.Unindent();
            Logger.Log("Block exit", 3);
        }

        #region IDisposable Support
        bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            LogBlockEnd();
            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~BlockLogger()
        {
            Dispose(false);
        }

        #endregion
    }
}