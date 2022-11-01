using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AccessCodeLib.AccUnit.AccessTestClientTests
{
    internal class TestResultCollector : ITestResultCollector
    {
        private readonly List<ITestResult> _results = new List<ITestResult>();

        public void Add(ITestResult testResult)
        {
            _results.Add(testResult);
        }

        public IEnumerable<ITestResult> Results => _results;
    }
}
