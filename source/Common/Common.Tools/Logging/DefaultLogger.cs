using System;
using System.Diagnostics;

namespace AccessCodeLib.Common.Tools.Logging
{
    public class DefaultLogger : LoggerBase
    {
        private const char _indentChar = ' ';
        private int IndentLevel { get; set; }
        private const int _indentSize = 1;

        private int IndentSize => _indentSize;

        public override void Log(string info, int addNumberStackFrames)
        {
            var message = $"{GetTimingInfo()}: {GetIndent()}{GetContextInfo(addNumberStackFrames)}";
            if (!string.IsNullOrEmpty(info))
            {
                message += ": " + info;
            }

            Debug.WriteLine(message);

            RememberPointInTime();
        }

        private string GetIndent()
        {
            return new string(IndentChar, IndentLevel * IndentSize);
        }

        private static char IndentChar => _indentChar;

        public override void LogRaw(string rawInfo)
        {
            Debug.WriteLine(rawInfo);
        }

        public override void Log(Exception exception, int addNumberStackFrames)
        {
            Log(exception.Message, addNumberStackFrames);
            LogRaw(exception.StackTrace);
        }

        public override void Indent()
        {
            IndentLevel++;
        }

        public override void Unindent()
        {
            IndentLevel--;
            if (IndentLevel >= 0) return;
            Debug.WriteLine("************** IndentLevel got negative! ********");
            IndentLevel = 0;
        }

        
    }
}