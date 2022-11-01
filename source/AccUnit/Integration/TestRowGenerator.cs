using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using AccessCodeLib.Common.VBIDETools;
using AccessCodeLib.Common.VBIDETools.TypeLib;
using Microsoft.Vbe.Interop;
using TLI;

namespace AccessCodeLib.AccUnit
{
    public class TestRowGenerator
    {
        public string TestName { get; set; }
        public VBProject ActiveVBProject { get; set; }

        /*
        public void GetTestData(TestDataBuilder databuilder)
        {
            GetTestData(databuilder, null);
        }

        public void GetTestData(_TestDataBuilder databuilder, TestClassMemberInfo memberinfo)
        {
            var rows = GetTestRows(databuilder.MethodName);
            var testRowFilter = (memberinfo != null ? memberinfo.TestRowFilter : null);

            if (testRowFilter != null && testRowFilter.Count > 0)
            {
                foreach (var rowindex in testRowFilter)
                {
                    try
                    {
                        var row = rows[rowindex];
                        CallDataBuilderUse(databuilder, row);
                    }
                    catch (IndexOutOfRangeException ex) // row deleted in source code?
                    {
                        Logger.Log(string.Format("TestRowGenerator AddRow [rowindex {0}]: {1}", rowindex, ex.Message));
                    }
                }
            }
            else
            {
                foreach (var row in rows)
                {
                    CallDataBuilderUse(databuilder, row);
                }
            }

            if (memberinfo != null)
                memberinfo.TestRows.AddRange(rows);

        }

        private static void CallDataBuilderUse(_TestDataBuilder databuilder, ITestRow row)
        {
            var arg = row.Args[0];
            var arrayLength = row.Args.Count - 1;
            var args = new object[arrayLength];
            Array.Copy(row.Args.ToArray(), 1, args, 0, arrayLength);
            databuilder.Use(ref arg, args).TestName(GetTestFixtureRowName(row));
        }
        */

        private static string GetTestFixtureRowName(ITestRow row)
        {
            //return row.Name != null ? string.Format("Row{0}{1} {2}", row.Index + 1, RowNameDelimiter, row.Name) : string.Format("Row{0}", row.Index + 1);
            return row.TestFixtureRowName;
        }

        private Queue<ITestRow> _rowCollection = new Queue<ITestRow>();

        public ITestRow AddRow(params object[] args)
        {
            ITestRow row = new TestRow(args);
            _rowCollection.Enqueue(row);
            row.Index = _rowCollection.Count - 1;
            Logger.Log(string.Format("Row index: {0}", row.Index));
            return row;
        }

        public ITestRow[] GetTestRows(string methodName)
        {
            using (new BlockLogger())
            {
                using (var m = new CodeModuleContainer(ActiveVBProject))
                {
                    var codeModuleReader = m.GetCodeModulReader(TestName);
                    var procHeader = codeModuleReader.GetProcedureHeader(methodName);
                    return GetRowsFromProcHeader(procHeader, methodName);
                }
            }
        }

        internal ITestRow[] GetRowsFromProcHeader(string procheader, string procname)
        {
            using (new BlockLogger())
            {
                _rowCollection = new Queue<ITestRow>();
                var testParams = GetRowTestParamStrings(procheader, procname);
                Logger.Log("next step: RunTestRowGeneratorCode");
                RunTestRowGeneratorCode(testParams, this);
                Logger.Log(string.Format("Rows erfasst: {0}", _rowCollection.Count));
                return _rowCollection.ToArray();
            }
        }

        // ReSharper disable ReturnTypeCanBeEnumerable.Local
        internal string[] GetRowTestParamStrings(string procHeader, string procname)
        // ReSharper restore ReturnTypeCanBeEnumerable.Local
        {
            procHeader = procHeader.Replace("\r", "");
            const string commentExtenstion = @"(\s*|\s*\'.*)";

            var pattern = string.Format(@"^\s*\'\s*(AccUnit:Row|TestManager\.Row|{0})(\(.*\).*){1}$", procname, commentExtenstion);
            var convertProcHeaderLineRegex = new Regex(pattern,
                                                       RegexOptions.CultureInvariant | RegexOptions.Multiline | RegexOptions.IgnoreCase);
            var testparams = new List<string>();
            foreach (var paramString in from Match m in convertProcHeaderLineRegex.Matches(procHeader)
                                        let parameterList = GetCheckedParameterList(m.Groups[2].Value)
                                        let optionalComment = m.Groups[3].Value
                                        select ConvertVbaParamStringToVB(parameterList) + optionalComment)
            {
                testparams.Add(paramString);
                Logger.Log(paramString);
            }
            return testparams.ToArray();
        }

        private static readonly Regex ParamStringNameProcedureRegex = new Regex(@"\.Name\(", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static readonly Regex ParamStringNamePropertyRegex = new Regex(@"(.*)Name\s*=\s*(.*)", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private static string GetCheckedParameterList(string paramstring)
        {
            paramstring = ParamStringNameProcedureRegex.Replace(paramstring, ".SetName("); //issue 81

            var match = ParamStringNamePropertyRegex.Match(paramstring);
            if (match.Success)
            {
                var nameString = match.Groups[2].Value;
                if (nameString.IndexOf("\"", StringComparison.Ordinal) != 0) 
                    nameString = "\"" + nameString.TrimEnd();

                if (nameString.LastIndexOf("\"", StringComparison.Ordinal) == 0)
                {
                    var commentStartIndex = nameString.LastIndexOf("'", StringComparison.Ordinal);

                    if (commentStartIndex > 0)
                    {
                        nameString = nameString.Substring(0, commentStartIndex - 1) + "\" " + nameString.Substring(commentStartIndex);
                    }
                    else
                        nameString = nameString + "\"";
                }

                paramstring = match.Groups[1] + "Name=" + nameString;
            }
            return paramstring;
        }

        private static readonly Regex ParamStringDbNullRegex = new Regex(@"([\(\,]?)\s*(Null)\s*([\)\,])", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private string ConvertVbaParamStringToVB(string paramstring)
        {
            var tempString = ParamStringDbNullRegex.Replace(paramstring,
                                                               m =>
                                                               string.Format("{0}{1}{2}", m.Groups[1].Value,
                                                                             "DBNull.Value", m.Groups[3].Value));
            tempString = ConvertConstantStringsToVB(tempString);
            return "TestManager.AddRow" + tempString;
        }

        private static readonly Regex ConstantStringRegex = new Regex(@"([\(\,]?)\s*([A-z\.]+)\s*([\)\,])", RegexOptions.CultureInvariant | RegexOptions.Compiled | RegexOptions.Multiline | RegexOptions.IgnoreCase);
        private string ConvertConstantStringsToVB(string paramstring)
        {
            var constants = ConstantsReader.Constants;
            if (constants.Count == 0 && ActiveVBProject != null)
                ConstantsReader.AddConstants(ActiveVBProject);
            
            Logger.Log(string.Format("Fill params, Constants: {0}", constants.Count));

            if (constants.Count == 0)
            {
                Logger.Log("Missing constants!");
                return paramstring;
            }

            var tempString = ConstantStringRegex.Replace(paramstring,
                                                            m => 
                                                            string.Format("{0}{1}{2}", m.Groups[1].Value,
                                                                             ReplaceReplaceVbConstant(m.Groups[2].Value), m.Groups[3].Value));
            Logger.Log("completed");
            return tempString;
        }

        private static string ReplaceReplaceVbConstant(string paramstring)
        {
            Logger.Log(string.Format("input: >{0}.<", paramstring));

            var parts = paramstring.Split('.');
            string searchString;

            switch (parts.Length)
            {
                case 1:
                    searchString = string.Format("{0}.", parts[0]);
                    break;
                case 2:
                    searchString = string.Format("{0}.{1}.", parts[1], parts[0]);
                    break;
                case 3:
                    searchString = string.Format("{0}.{1}.{2}", parts[2], parts[1], parts[0]);
                    break;
                default:
                    return paramstring;
            }

            Constant constant;
            try
            {
                if (parts.Length == 3)
                    constant = ConstantsReader.Constants[searchString];
                else
                    constant =
                        ConstantsReader.Constants.First(
                            c => c.Key.StartsWith(searchString, StringComparison.CurrentCultureIgnoreCase)).Value;
                
            }
            catch
            {
                return paramstring;
            }

            if (constant==null)
                return paramstring;
            

            string valueString;
            switch (constant.VarType)
            {
                case TliVarType.VT_BSTR:
                    valueString = string.Format("\"{0}\"", constant.Value);
                    break;
                default:
                    valueString = constant.Value.ToString();
                    break;
            }

            return valueString;
        }


        private static object CreateTestRowGenerator(string testparamstring)
        {
            var sourcecode = GetTestRowGeneratorSource(testparamstring);

            //Logger.Log(sourcecode);

            using (var bcp = new Microsoft.VisualBasic.VBCodeProvider())
            {
                var cp = new CompilerParameters();
                var results = bcp.CompileAssemblyFromSource(cp, sourcecode);
                EnsureCouldCompile(results);
                return results.CompiledAssembly.CreateInstance("DynamicTestRowGenerator");
            }

        }

        private static void EnsureCouldCompile(CompilerResults results)
        {
            if (results.Errors.HasErrors)
            {
                throw new CouldNotCompileDynamicTestRowGeneratorException(results);
            }
        }

        private static string GetTestRowGeneratorSource(string rowtestcode)
        {
            return @"
Imports System
Imports System.Reflection
Public class DynamicTestRowGenerator
  Public Sub InsertTestData(TestManager As Object)
      " + rowtestcode + @"
  End Sub
End Class";
        }

        private static void RunTestRowGeneratorCode(IEnumerable<string> testparams, TestRowGenerator host)
        {
            using (new BlockLogger())
            {
                var testRowLines = GetTestRowLines(testparams);
                //Logger.Log(testRowLines);
                var rowGenerator = CreateTestRowGenerator(testRowLines);
                //Logger.Log(rowGenerator.ToString());
                InvokeInsertTestDataProvidingInnerException(rowGenerator, host);
            }
        }

        private static void InvokeInsertTestDataProvidingInnerException(object rowGenerator, TestRowGenerator host)
        {
            try
            {
                rowGenerator.GetType().GetMethod("InsertTestData").Invoke(rowGenerator, new object[] { host });
            }
            catch (TargetInvocationException xcp)
            {
                throw xcp.InnerException;
            }
        }

        private static string GetTestRowLines(IEnumerable<string> testparams)
        {
            return testparams.Aggregate<string, string>(null, (current, paramString) => current + string.Format("{0}\r\n", paramString));
        }


    }
}
