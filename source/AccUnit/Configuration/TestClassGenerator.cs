using AccessCodeLib.Common.VBIDETools;
using Microsoft.Vbe.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace AccessCodeLib.AccUnit.Configuration
{

    [ComVisible(true)]
    [Guid("F308DAC7-4CD0-4C37-B085-AF436D731034")]
    public interface ITestClassGenerator
    {
        void NewTestClass(string ClassToTest = null);
        void EditTestClass(string TestClassName = null);
    }

    [ComVisible(true)]
    [Guid("333BD1B7-23BD-44E8-833D-E11627108223")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("AccUnit.TestClassGenerator")]
    public class TestClassGenerator : ITestClassGenerator
    {
        public TestClassGenerator()
        {
        }

        public TestClassGenerator(OfficeApplicationHelper applicationHelper)
        {
            ApplicationHelper = applicationHelper;
        }

        public VBProject ActiveVBProject { get { return ApplicationHelper.CurrentVBProject; } }
        public OfficeApplicationHelper ApplicationHelper { get; set; }

        public void EditTestClass(string TestClassName = null)
        {
            throw new NotImplementedException();
        }

        public void NewTestClass(string ClassToTest = null)
        {
            throw new NotImplementedException();
        }
    }
}
