using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using AccessCodeLib.Common.Tools.Logging;

namespace AccessCodeLib.Common.VBIDETools
{
    public class CodeModuleContainer : IDisposable
    {
        private VBProject _vbProject;
        private readonly CodeModuleGenerator _codeModuleGenerator;

        public CodeModuleContainer(VBProject vbProjectToUse)
        {
            _vbProject = vbProjectToUse;
            _codeModuleGenerator = new CodeModuleGenerator(vbProjectToUse);
        }

        public CodeModuleGenerator Generator { get { return _codeModuleGenerator; } }

        public CodeModuleReader GetCodeModulReader(string name)
        {
            var cm = GetCodeModule(name);
            return new CodeModuleReader(cm);
        }

        private CodeModule GetCodeModule(string name)
        {
            var c = _vbProject.VBComponents.Item(name);
            return c.CodeModule;
        }

        public CodeModule TryGetCodeModule(string name)
        {
            try
            {
                return GetCodeModule(name);
            }
            catch (IndexOutOfRangeException)
            {
                return null;
            }
        }

        public bool Exists(string name)
        {
            return (TryGetCodeModule(name) != null);
        }

        public void Remove(string name)
        {
            foreach (var vbc in
                _vbProject.VBComponents.Cast<VBComponent>().Where(vbc => vbc.Name.Equals(name, StringComparison.InvariantCultureIgnoreCase)))
            {
                Remove(vbc);
                break;
            }
        }

        public void Remove(VBComponent component)
        {
            _vbProject.VBComponents.Remove(component);
        }

        public string Export(string name, string exportfolder)
        {
            var c = _vbProject.VBComponents.Item(name);
            return Export(c, exportfolder);
        }

        public string Export(VBComponent component, string exportfolder)
        {
            var fileName = $"{exportfolder.TrimEnd(' ', '\\')}\\{component.Name}{GetVbComponentTypeFileExtension(component.Type)}";
            component.Export(fileName);
            return fileName;
        }

        private static string GetVbComponentTypeFileExtension(vbext_ComponentType componenttype)
        {
            string fileExtension;
            switch (componenttype)
            {
                case vbext_ComponentType.vbext_ct_ClassModule:
                    fileExtension = ".cls";
                    break;
                case vbext_ComponentType.vbext_ct_StdModule:
                    fileExtension = ".bas";
                    break;
                case vbext_ComponentType.vbext_ct_MSForm:
                    fileExtension = ".frm";
                    break;
                case vbext_ComponentType.vbext_ct_ActiveXDesigner:
                    fileExtension = ".dsr";
                    break;
                default:
                    fileExtension = ".acm";
                    break;
            }
            return fileExtension;
        }

        public string ExportAndRemove(string name, string exportfolder)
        {
            var c = _vbProject.VBComponents.Item(name);
            return ExportAndRemove(c, exportfolder);
        }

        public string ExportAndRemove(VBComponent component, string exportfolder)
        {
            var fileName = Export(component, exportfolder);
            Remove(component);
            return fileName;
        }

        #region IDisposable Support

        bool _disposed;
        protected void Dispose(bool disposing)
        {
            if (_disposed) return;

            try
            {
                if (disposing)
                    DisposeManagedResources();

                DisposeUnmanagedResources();
                _disposed = true;
            }
            catch (Exception ex)
            {
                Logger.Log(ex.Message);
            }
        }

        void DisposeManagedResources()
        {
            _codeModuleGenerator.Dispose();
        }

        void DisposeUnmanagedResources()
        {
            _vbProject = null;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~CodeModuleContainer()
        {
            Dispose(false);
        }

        #endregion
    }
}
