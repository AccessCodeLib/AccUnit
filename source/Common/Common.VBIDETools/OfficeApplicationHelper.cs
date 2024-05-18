using AccessCodeLib.Common.Tools.Files;
using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using System;
using System.Linq;
using System.Runtime.InteropServices;

namespace AccessCodeLib.Common.VBIDETools
{
    public interface IOfficeApplicationHelper
    {
        object Application { get; }
        VBProject CurrentVBProject { get; }
        string Name { get; }
        VBE VBE { get; }

        void Dispose();
        object Run(params object[] parameters);
    }

    public class OfficeApplicationHelper : IOfficeApplicationHelper, IDisposable
    {
        private object _application;
        private InvocationHelper _invocationHelper;

        public object Application { get { return _application; } }
        protected InvocationHelper InvocationHelper { get { return _invocationHelper; } }

        private string _officeApplicationName;
        public string Name
        {
            get
            {
                if (string.IsNullOrEmpty(_officeApplicationName))
                    _officeApplicationName = (string)_invocationHelper.InvokePropertyGet("Name");

                return _officeApplicationName;
            }
        }

        public OfficeApplicationHelper(object application)
        {
            _application = application ?? throw new ArgumentNullException();
            _invocationHelper = new InvocationHelper(application);
        }

        public virtual VBE VBE
        {
            get { return (VBE)_invocationHelper.InvokePropertyGet("VBE"); }
        }

        public VBProject CurrentVBProject
        {
            get
            {
                var checkedVbProject = GetCheckedVbProject();
                return checkedVbProject ?? VBE.ActiveVBProject;
            }
        }

        /// @todo check ActiveVbProject
        protected virtual VBProject GetCheckedVbProject()
        {
            var activeVBProject = VBE.ActiveVBProject;
            try
            {
                var activeDocumentFullName = FileTools.ConvertPathToUNC(GetDocumentFullName());
                if (string.IsNullOrEmpty(activeDocumentFullName) && !VbProjectMatchWithFullName(activeVBProject, activeDocumentFullName))
                {
                    activeVBProject = FindVbProjectEqualsFullName(activeDocumentFullName);
                }
                return activeVBProject;
            }
            catch (Exception ex)
            {
                Logger.Log(ex);
                return null;
            }
        }

        private string GetDocumentFullName()
        {
            // TODO: The strings are not sufficient to be able to invoke the infos. We also need the kind of invocation (method or getter call).
            //       For example, the member CurrentDb in Access.Application is a method, not a getter.
            var documentInvocatioHelper = new InvocationHelper(_invocationHelper.InvokeMethod(OfficeApplcationInvokeStrings.ActiveDocument));
            return (string)documentInvocatioHelper.InvokePropertyGet(OfficeApplcationInvokeStrings.ActiveDocumentInvokeStrings.FullName);
        }

        protected static bool VbProjectMatchWithFullName(_VBProject vbProjectToTest, string fullName)
        {
            return fullName.Equals(TryGetFileNameFromVbProjectAndIngoreErrors(vbProjectToTest), StringComparison.InvariantCultureIgnoreCase);
        }

        private static string TryGetFileNameFromVbProjectAndIngoreErrors(_VBProject vbProject)
        {
            try
            {
                return vbProject.FileName;
            }
            catch (Exception)
            {
                return string.Empty;
            }
        }

        protected VBProject FindVbProjectEqualsFullName(string fullName)
        {
            return VBE.VBProjects.Cast<VBProject>().First(vbProject => VbProjectMatchWithFullName(vbProject, fullName));
        }

        private IOfficeApplicationInvokeStrings _officeApplcationInvokeStrings;
        private IOfficeApplicationInvokeStrings OfficeApplcationInvokeStrings
        {
            get { return _officeApplcationInvokeStrings ?? (_officeApplcationInvokeStrings = OfficeInvokeNamesBuilder.OfficeApplicationInvokeStrings(Name)); }
        }

        public object Run(params object[] parameters)
        {
            var result = InvocationHelper.InvokeMethod(OfficeApplcationInvokeStrings.Run, parameters);
            return result;
        }

        #region IDisposable Support

        bool _disposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_disposed) return;

            try
            {
                if (disposing)
                {
                    if (_invocationHelper != null)
                    {
                        _invocationHelper.Dispose();
                        _invocationHelper = null;
                    }
                }

                if (_application != null)
                {
                    Marshal.ReleaseComObject(_application); 
                    _application = null;
                }   
                
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }

            // GC-Bereinigung wegen unmanaged res:
            /*
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            */
            _disposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~OfficeApplicationHelper()
        {
            Dispose(false);
        }

        #endregion

    }
}