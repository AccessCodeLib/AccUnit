﻿using AccessCodeLib.Common.VBIDETools;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class CheckableCodeModuleInfo : CheckableItem
    {
        private CodeModuleInfo _codeModule; 
        public CheckableCodeModuleInfo(CodeModuleInfo codeModule, bool isChecked = false) : base(codeModule.Name, isChecked)
        {
            _codeModule = codeModule;
        }

        public CodeModuleInfo CodeModule { get { return _codeModule; } }    
    }
}
