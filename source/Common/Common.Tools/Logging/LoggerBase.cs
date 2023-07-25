using System;
using System.Diagnostics;

namespace AccessCodeLib.Common.Tools.Logging
{
    public abstract class LoggerBase : ILogger
    {
        private readonly AdaptiveFormatter _timingFormatter = new AdaptiveFormatter("{{0,{0}:#,##0.0}}ms") { CurrentInfoLength = 8 };
        private readonly AdaptiveFormatter _contextFormatter = new AdaptiveFormatter("{{0,-{0}}}") { CurrentInfoLength = 60 };

        private static DateTime? PrecedingEventTimestamp { get; set; }

        private AdaptiveFormatter TimingFormatter
        {
            get { return _timingFormatter; }
        }

        private AdaptiveFormatter ContextFormatter
        {
            get { return _contextFormatter; }
        }

        protected string GetContextInfo(int addNumberStackFrames)
        {
            return ContextFormatter.GetFormattedInfo(GetContextInfoInternal(addNumberStackFrames));
        }

        private static string GetContextInfoInternal(int addNumberStackFrames)
        {
            var stackFrame = new StackTrace().GetFrame(4 + addNumberStackFrames);
            var method = stackFrame.GetMethod();
            var methodName = method.Name;
            var typeName = method.DeclaringType.Name;

            return string.Format("{0}.{1}", typeName, methodName);
        }

        protected string GetTimingInfo()
        {
            if (PrecedingEventTimestamp.HasValue)
            {
                var elapsedTime = DateTime.Now - PrecedingEventTimestamp;
                var elapsedMilliseconds = elapsedTime.Value.TotalMilliseconds;

                return TimingFormatter.GetFormattedInfo(elapsedMilliseconds);
            }
            return string.Format(new string('-', TimingFormatter.CurrentInfoLength + "ms".Length));
        }


        protected static void RememberPointInTime()
        {
            PrecedingEventTimestamp = DateTime.Now;
        }

        public abstract void Log(string info, int addNumberStackFrames);
        public abstract void LogRaw(string rawInfo);
        public abstract void Log(Exception exception, int addNumberStackFrames);
        public abstract void Indent();
        public abstract void Unindent();
    }
}