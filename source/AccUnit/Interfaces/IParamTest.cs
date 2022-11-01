using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.Interfaces
{
    public interface IParamTest : ITest
    {
        IEnumerable<object> Parameters { get; }
    }    
 }
