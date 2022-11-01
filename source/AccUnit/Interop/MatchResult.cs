using System.Runtime.InteropServices;
using System;

namespace AccessCodeLib.AccUnit.Interop
{
    /*
    [ComVisible(true)]
    [Guid("1E679EDC-C204-48B6-9630-0AE1AF8DB290")]
    public interface IMatchResult : AccUnit.Assertions.IMatchResult
    {
        new bool Match { get; }
        new string Text { get; } // [return: MarshalAs(UnmanagedType.LPWStr)] 
        new string Actual { get; }
        new string Expected { get; }
    }
    */
    
    [ComVisible(true)]
    [Guid("6AB715D1-A0A7-4310-91BA-25921165E716")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".MatchResult")]
    public class MatchResult : AccUnit.Assertions.MatchResult, Assertions.IMatchResult
    {
        public MatchResult(bool match, string text, object actual, object expected, string infoText = null) 
            : base(null, match, text, actual, expected, infoText)
        {
        }

        public MatchResult(AccUnit.Assertions.IMatchResult result) 
            : base(null, result.Match, result.Text, result.Actual, result.Expected, result.InfoText)
        {
            //Match = result.Match;
            //Text = result.Text;
        }

        new public bool Match { get { return base.Match; } }
        
        new public string Actual { get { return convertToString(base.Actual); } }
        new public string Expected { get { return convertToString(base.Expected); } }

        new public string Text { get { return base.Text; } }
   
    }
}
