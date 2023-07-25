using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using System;
using System.Linq;

namespace AccessCodeLib.Common.VBIDETools
{
    public class CodeModuleGenerator : IDisposable
    {
        private VBProject _vbProject;

        public CodeModuleGenerator(VBProject activeVBProject)
        {
            _vbProject = activeVBProject;
        }

        public CodeModule Add(System.IO.FileInfo fileInfo)
        {
            return _vbProject.VBComponents.Import(fileInfo.FullName).CodeModule;
        }

        public CodeModule Add(vbext_ComponentType type,
                              string name,
                              string sourcecode = null,
                              bool removeStandardLines = false)
        {
            var vbcomponent = AddComponent(name, type);
            var codemodule = vbcomponent.CodeModule;

            if (removeStandardLines)
                codemodule.DeleteLines(1, codemodule.CountOfLines);

            if (sourcecode != null)
                codemodule.InsertLines(1, sourcecode);

            return codemodule;
        }

        private VBComponent AddComponent(string moduleName, vbext_ComponentType componentType)
        {
            EnsureModuleNameIsNotInUse(moduleName);
            try
            {
                var vbc = _vbProject.VBComponents.Add(componentType);
                vbc.Name = moduleName;
                return vbc;
            }
            catch (Exception xcp)
            {
                var message = string.Format("Could not add VBComponent \"{0}\".\nMaybe the office file is write protected or the user has insufficient privileges.", moduleName);
                throw new Exception(message, xcp);
            }
        }

        private void EnsureModuleNameIsNotInUse(string moduleName)
        {
            const string messageTemplate = "Cannot add module with the name \"{0}\" to the VBProject. There is already a module with that name.";

            if (GetModuleOrNull(moduleName) != null)
            {
                throw new ArgumentException(string.Format(messageTemplate, moduleName));
            }
        }

        private VBComponent GetModuleOrNull(string moduleName)
        {
            return _vbProject.VBComponents.Cast<VBComponent>().FirstOrDefault(component => component.Name == moduleName);
        }


        #region IDisposable Support

        bool _disposed;
        protected void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            try
            {
                //if (disposing)
                //{
                //    DisposeManagedResources();
                //}
                DisposeUnmanagedResources();
                _disposed = true;
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }
        }

        //void DisposeManagedResources()
        //{
        //}

        void DisposeUnmanagedResources()
        {
            _vbProject = null;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~CodeModuleGenerator()
        {
            Dispose(false);
        }

        #endregion
    }
}
