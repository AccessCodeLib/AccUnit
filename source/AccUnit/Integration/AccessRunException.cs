using System;
using System.Collections.Generic;

namespace AccessCodeLib.AccUnit
{
    public class OfficeApplicationRunException : Exception
    {
        public OfficeApplicationRunException()
            : this("An exception occurred while calling Application.Run()")
        {
        }

        public OfficeApplicationRunException(string message)
            : base(message)
        {
            _parameters = new object[0];
        }

        public OfficeApplicationRunException(string message, Exception innerException)
            : base(message, innerException)
        {
            _parameters = new object[0];
        }

        public OfficeApplicationRunException(Exception innerException, object[] parameters)
            : base(
                string.Format("An exception occurred while calling method {0} via Application.Run()", parameters?[0]),
                innerException)
        {
            if (parameters != null && parameters.Length > 0)
            {
                MethodName = Convert.ToString(parameters[0]);
                _parameters = new object[parameters.Length - 1];
                Array.Copy(parameters, 1,
                           _parameters, 0,
                           parameters.Length - 1);
            }
            else
            {
                _parameters = new object[0];
            }
        }

        private readonly object[] _parameters;
        public IList<object> Parameters { get { return _parameters; } }

        public string MethodName { get; private set; }
    }
}