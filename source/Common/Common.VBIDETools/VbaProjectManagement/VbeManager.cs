using AccessCodeLib.Common.Tools.Logging;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AccessCodeLib.Common.VBIDETools.VbaProjectManagement
{
    public class VbeManager : IVbeManager
    {
        private readonly VBProject _vbProject;
        private IEnumerable<Module> _modules;

        public VbeManager(VBProject vbProject)
        {
            _vbProject = vbProject ?? throw new ArgumentNullException("vbProject");
            Logger.Log(string.Format("Name of VBProject: \"{0}\"", _vbProject.Name));
        }

        public void WriteOrCreate(Module module)
        {
            var vbComponent = _vbProject.VBComponents.Cast<VBComponent>()
                .Where(vbc => vbc.Name == module.Name)
                .SingleOrDefault();
            if (vbComponent is null)
            {
                vbComponent = _vbProject.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
                vbComponent.Name = module.Name;
            }
            // TODO: Make this more generic/configurable
            var codeModule = vbComponent.CodeModule;
            codeModule.DeleteLines(1, codeModule.CountOfLines);
            codeModule.AddFromString("Option Compare Text");
            codeModule.AddFromString("Option Explicit");
            codeModule.AddFromString("");
            codeModule.AddFromString(module.GetNewContent());
        }

        public string ProjectName
        {
            get { return _vbProject.Name; }
        }

        public IEnumerable<Module> Modules
        {
            get
            {
                if (_modules is null)
                {
                    _modules = GetModules();
                }
                return _modules;
            }
        }

        private IEnumerable<Module> GetModules()
        {
            using (new BlockLogger())
            {
                return _vbProject.VBComponents
                    .Cast<VBComponent>()
                    .Select(vbc => new Module(vbc.Name, () => vbc.CodeModule));
            }
        }
    }
}