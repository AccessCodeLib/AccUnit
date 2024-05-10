using AccessCodeLib.AccUnit.Interfaces;
using AccessCodeLib.Common.Tools.Logging;
using System;
using System.Collections.Generic;
using VbMsgBoxResult = AccessCodeLib.AccUnit.Interfaces.VbMsgBoxResult;

namespace AccessCodeLib.AccUnit
{
    public class TestRow : ITestRow, _ITestRow
    {
        public const string RowNameDelimiter = ": ";

        public TestRow(object[] args)
        {
            if (args.Length < 1)
            {
                throw new ArgumentOutOfRangeException("args", @"The datarow must contain at least one parameter.");
            }
            _args = args.Length > 0 ? args : new object[] { };
        }

        private readonly object[] _args;
        public IList<object> Args { get { return _args; } }

        public int Index { get; set; }

        public string Name { get; set; }

        public ITestRow SetName(string name)
        {
            using (new BlockLogger("TestRow.SetName"))
            {
                Name = name;
                Logger.Log(string.Format("Name: {0}", name));
            }
            return this;
        }

        public string TestFixtureRowName
        {
            get
            {
                return string.IsNullOrEmpty(Name)
                           ? string.Format("Row{0}", Index + 1)
                           : string.Format("Row{0}{1} {2}", Index + 1, RowNameDelimiter, Name);
            }
        }

        public ITestMessageBox TestMessageBox { get; set; }

        public ITestRow ClickingMsgBox(params VbMsgBoxResult[] args)
        {
            using (new BlockLogger("TestRow.ClickingMsgBox"))
            {
                try
                {
                    TestMessageBox = new TestMessageBox();
                    Logger.Log(string.Format("Anzahl Args: {0}", args.Length));
                    TestMessageBox.InitMsgBoxResults(args);
                }
                catch (Exception ex)
                {
                    Logger.Log(ex);
                }

            }
            return this;
        }

        private readonly IgnoreInfo _ignoreInfo = new IgnoreInfo();
        public IgnoreInfo IgnoreInfo { get { return _ignoreInfo; } }

        public ITestRow Ignore(string comment = "")
        {
            using (new BlockLogger("TestRow.Ignore"))
            {
                _ignoreInfo.Ignore = true;
                _ignoreInfo.Comment = comment;
                Logger.Log(string.Format("Ignore=True, Comment: {0}", comment));
            }
            return this;
        }

        private readonly ITagList _tags = new TagList();
        public ITagList Tags { get { return _tags; } }

        public ITestRow AddTags(params string[] args)
        {
            foreach (var arg in args)
            {
                (_tags as TagList).Add(new TestItemTag(arg));
            }
            return this;
        }

    }
}