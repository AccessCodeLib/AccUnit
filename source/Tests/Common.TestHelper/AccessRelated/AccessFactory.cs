using System;

namespace AccessCodeLib.Common.TestHelpers.AccessRelated
{
    static class AccessFactory
    {
        private const string ProgId = "Access.Application";

        public static dynamic CreateApplication()
        {
            var type = Type.GetTypeFromProgID(ProgId);
            if (type == null)
                throw new Exception(string.Format("Could not locate {0}.", ProgId));

            var instance = Activator.CreateInstance(type);
            if (instance == null)
                throw new Exception(string.Format("Error on creating an instance of {0}.", ProgId));

            return instance;
        }
    }
}