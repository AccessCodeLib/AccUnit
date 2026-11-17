using System;
using System.Collections.Generic;
using System.Linq;
using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;

namespace AccessCodeLib.AccUnit.McpServer
{
    /// <summary>
    /// Selects the requested test classes and test related modules for export or removal; other modules are rejected,
    /// so these tools never touch production code.
    /// </summary>
    public static class TestModuleSelection
    {
        private const string ClassModuleExtension = ".cls";
        private const string StandardModuleExtension = ".bas";

        /// <param name="requestedNames">Module names from the client.</param>
        /// <param name="testModules">Test classes and test related modules of the VBA project.</param>
        /// <param name="componentNames">All modules of the VBA project.</param>
        /// <param name="rejected">Receives the requested names that are not selected, with the reason.</param>
        public static List<CodeModuleInfo> Select(IList<string> requestedNames, IEnumerable<CodeModuleInfo> testModules,
                                                  ICollection<string> componentNames, List<RejectedTestModule> rejected)
        {
            if (requestedNames == null || requestedNames.Count == 0 || requestedNames.Any(string.IsNullOrWhiteSpace))
                throw new DatabaseSessionException("Name the test classes (classNames); empty names are not allowed.");

            var testModuleList = testModules.ToList();
            var selected = new List<CodeModuleInfo>();
            foreach (var name in requestedNames.Select(n => n.Trim()).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var testModule = testModuleList.FirstOrDefault(m => string.Equals(m.Name, name, StringComparison.OrdinalIgnoreCase));
                var reason = GetRejectReason(name, testModule, componentNames);
                if (reason == null)
                    selected.Add(testModule);
                else
                    rejected.Add(new RejectedTestModule { Name = name, Reason = reason });
            }
            return selected;
        }

        private static string GetRejectReason(string name, CodeModuleInfo testModule, ICollection<string> componentNames)
        {
            if (testModule == null)
                return componentNames.Contains(name, StringComparer.OrdinalIgnoreCase)
                    ? "Not a test module: the header has no comment 'AccUnit:TestClass or 'AccUnit:TestRelated."
                    : "No module with this name.";
            if (GetFileExtension(testModule.ComponentType) == null)
                return "Only class modules and standard modules are supported, not form or report modules.";
            return null;
        }

        /// <returns>File extension of the VBA export, or null for module types the tools do not support.</returns>
        public static string GetFileExtension(vbext_ComponentType componentType)
        {
            switch (componentType)
            {
                case vbext_ComponentType.vbext_ct_ClassModule: return ClassModuleExtension;
                case vbext_ComponentType.vbext_ct_StdModule: return StandardModuleExtension;
                default: return null;
            }
        }
    }

    public class RejectedTestModule
    {
        public string Name { get; set; }
        public string Reason { get; set; }
    }
}
