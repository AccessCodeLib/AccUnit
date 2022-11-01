using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Assertions
{
    [ComVisible(true)]
    [Guid("F3860CD4-1D20-40D0-A1C9-53F181316232")]
    public interface IMatchResult
    { 
        bool Match { get; }
        string Text { get; }
        object Actual { get; }
        object Expected { get; }

        string FormattedText { get; }
        string InfoText { get; set; }

        [ComVisible(false)]
        string CompareText { get; }
       
    }
}