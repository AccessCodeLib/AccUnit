﻿using System.Collections.Generic;
using AccessCodeLib.Common.VBIDETools;

namespace AccessCodeLib.AccUnit.Tools
{
    public interface ITestClassGenerator
    {
        string ClassName { get; set; }

        void Add(CodeModuleMember codeModuleMember);
        void Add(List<CodeModuleMember> codeModuleMembers);

        string GenerateSourceCode(); // => z. B. mit TextGen

        // und/oder direkt ein CodeModul erzeugen bzw. ergänzen:
        void Save();
        void SaveAs(string className);
    }
}