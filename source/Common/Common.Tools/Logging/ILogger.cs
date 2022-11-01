using System;

namespace AccessCodeLib.Common.Tools.Logging
{
    public interface ILogger
    {
        void Log(string info, int addNumberStackFrames = 0);
        void LogRaw(string rawInfo);
        void Log(Exception exception, int addNumberStackFrames = 0);
        void Indent();
        void Unindent();
    }
}