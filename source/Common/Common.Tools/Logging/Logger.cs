using System;
using System.Diagnostics;

namespace AccessCodeLib.Common.Tools.Logging
{
    public static class Logger
    {
        private static ILogger _instance = new DefaultLogger();

        private static ILogger Instance
        {
            get { return _instance; }
            set { _instance = value; }
        }

        [Conditional("DEBUG")]
        public static void Log(string info, int addNumberStackFrames = 0)
        {
            Instance.Log(info, addNumberStackFrames);
        }

        [Conditional("DEBUG")]
        public static void Log()
        {
            Instance.Log(string.Empty);
        }

        [Conditional("DEBUG")]
        public static void LogRaw(string rawInfo)
        {
            Instance.LogRaw(rawInfo);
        }

        [Conditional("DEBUG")]
        public static void Log(Exception exception, int addNumberStackFrames = 0)
        {
            Instance.Log(exception, addNumberStackFrames + 1);
        }

        [Conditional("DEBUG")]
        public static void Indent()
        {
            Instance.Indent();
        }

        [Conditional("DEBUG")]
        public static void Unindent()
        {
            Instance.Unindent();
        }
    }
}