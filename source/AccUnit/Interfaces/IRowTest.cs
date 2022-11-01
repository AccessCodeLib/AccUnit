using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface IRowTest : ITest
    {
        IEnumerable<ITestRow> Rows { get; }
        IEnumerable<IParamTest> ParamTests { get; }
    }
 }
