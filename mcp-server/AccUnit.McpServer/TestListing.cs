using System.Collections.Generic;

namespace AccessCodeLib.AccUnit.McpServer
{
    public class TestClassListing
    {
        public string ClassName { get; set; }
        public List<string> Tags { get; set; } = new List<string>();
        public List<TestMethodListing> Tests { get; set; } = new List<TestMethodListing>();
    }

    public class TestMethodListing
    {
        public string Name { get; set; }
        public List<string> Tags { get; set; } = new List<string>();
        public bool Ignored { get; set; }
        public string IgnoreComment { get; set; }

        /// <summary>Number of data rows of a row test; 0 for a simple test.</summary>
        public int RowCount { get; set; }
    }
}
