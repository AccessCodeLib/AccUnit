using System.Collections.Generic;

namespace AccessCodeLib.Common.VBIDETools.VbaProjectManagement
{
    public interface IVbeManager
    {
        void WriteOrCreate(Module module);
        string ProjectName { get; }
        IEnumerable<Module> Modules { get; }
    }
}