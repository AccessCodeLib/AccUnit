using System;
using AccessCodeLib.AccUnit.Properties;

namespace AccessCodeLib.AccUnit
{
    class MissingTestMessageBoxResultsException : Exception
    {
        public MissingTestMessageBoxResultsException()
            : this(Resources.MissingTestMessageBoxResult)
        {
        }

        public MissingTestMessageBoxResultsException(string message)
            :base(message)
        {
        }

        public MissingTestMessageBoxResultsException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public MissingTestMessageBoxResultsException(Exception innerException)
            : base(
                Resources.MissingTestMessageBoxResult,
                innerException)
        {
        }
    }
}
