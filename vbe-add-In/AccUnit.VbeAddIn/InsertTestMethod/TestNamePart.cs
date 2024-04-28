namespace AccessCodeLib.AccUnit.VbeAddIn
{
    internal class TestNamePart : ITestNamePart
    {
        public TestNamePart(string name, string caption)
        {
            Name = name;
            Caption = caption;
        }

        public string Name { get; private set; }
        public string Caption { get; private set; }
        public string Value { get; set; }
    }

}