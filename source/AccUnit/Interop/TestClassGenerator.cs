using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{

    [ComVisible(true)]
    [Guid("F308DAC7-4CD0-4C37-B085-AF436D731034")]
    public interface ITestClassGenerator
    {
        ITestMethodGenerator NewTestClass(string CodeModuleToTest = null, bool GenerateTestMethodsFromCodeModuleToTest = false, string stateUnderTest = null, string expectedBehaviour = null);
        ITestMethodGenerator EditTestClass(string TestClassName);
    }

    [ComVisible(true)]
    [Guid("936EB789-12F7-4820-AFED-7BE985D65E01")]
    public interface ITestMethodGenerator
    {
        ITestMethodGenerator InsertTestMethods(object MethodNamesUnderTest, string stateUnderTest, string expectedBehaviour);
        ITestMethodGenerator InsertTestMethod(string MethodNameUnderTest, string stateUnderTest, string expectedBehaviour);
    }

    [ComVisible(true)]
    [Guid("333BD1B7-23BD-44E8-833D-E11627108223")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccUnit.TestClassGenerator")]
    public class TestClassGenerator : Tools.TestClassGenerator, ITestClassGenerator, ITestMethodGenerator
    {

        private string _testClassName;

        public TestClassGenerator(VBProject vbproject) : base(vbproject)
        {
        }

        public ITestMethodGenerator EditTestClass(string TestClassName)
        {
            _testClassName = TestClassName;
            return this;
        }

        public ITestMethodGenerator NewTestClass(string CodeModuleToTest = null, bool GenerateTestMethodsFromCodeModuleToTest = false, string stateUnderTest = null, string expectedBehaviour = null)
        {
            _testClassName = base.NewTestClass(CodeModuleToTest, true, GenerateTestMethodsFromCodeModuleToTest, stateUnderTest, expectedBehaviour).Name;
            return this;
        }

        public ITestMethodGenerator InsertTestMethod(string MethodNameUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            InsertTestMethods(_testClassName, new string[] { MethodNameUnderTest }, stateUnderTest, expectedBehaviour);
            return this;
        }

        public ITestMethodGenerator InsertTestMethods(object MethodNamesUnderTest, string stateUnderTest, string expectedBehaviour)
        {
            IEnumerable<string> methods;

            if (MethodNamesUnderTest.GetType().IsArray)
            {
                methods = ((object[])MethodNamesUnderTest).Select(x => x.ToString()).ToArray();
            }
            else // check if MethodNamesUnderTest is a string and convert it to string array + check for , and ; as separators
            {
                if (MethodNamesUnderTest.GetType() == typeof(string))
                {
                    var MethodNamesUnderTestString = MethodNamesUnderTest.ToString();
                    if (MethodNamesUnderTestString.Contains(",") || MethodNamesUnderTestString.Contains(";"))
                    {
                        var methodNamesString = MethodNamesUnderTestString.Replace(";", ",").Replace(" ", "");
                        methods = methodNamesString.Split(',').Select(x => x.Trim()).ToArray();
                    }
                    else
                    {
                        methods = new string[] { MethodNamesUnderTest.ToString() };
                    }
                }
                else
                {
                    throw new ArgumentException("MethodNamesUnderTest must be a string or an array");
                }
            }

            InsertTestMethods(_testClassName, methods, stateUnderTest, expectedBehaviour);
            return this;
        }
    }
}
