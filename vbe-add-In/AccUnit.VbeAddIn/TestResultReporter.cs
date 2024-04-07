using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public class TestResultReporter : ITestResultReporter
    {

        private ITestResultCollector _testResultCollector;  

        public ITestResultCollector TestResultCollector { get; set; }




    }
}
