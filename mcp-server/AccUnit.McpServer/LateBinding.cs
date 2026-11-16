using System.Reflection;
using System.Runtime.ExceptionServices;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>Late bound calls (IDispatch) on COM objects of the Access process.</summary>
    internal static class LateBinding
    {
        public static object GetProperty(object target, string name, params object[] args)
        {
            return Invoke(target, name, BindingFlags.GetProperty, args);
        }

        public static object CallMethod(object target, string name, params object[] args)
        {
            return Invoke(target, name, BindingFlags.InvokeMethod, args);
        }

        private static object Invoke(object target, string name, BindingFlags flags, object[] args)
        {
            try
            {
                return target.GetType().InvokeMember(name, flags, null, target, args);
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw;
            }
        }
    }
}
