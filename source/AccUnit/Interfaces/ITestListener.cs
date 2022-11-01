using System;
using System.Collections.Generic;
using System.Text;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface ITestListener
    {
        ITestSuite TestSuite { get; set; }
    }
}
