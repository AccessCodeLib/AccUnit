using AccessCodeLib.AccUnit.Assertions;
using AccessCodeLib.AccUnit.Interfaces;
using System;
using System.Runtime.InteropServices;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("0011016E-BF37-4CB7-9A62-58DD78292550")]
    public interface IConstraintBuilder : AccUnit.Assertions.IConstraintBuilder 
    {

        IConstraintBuilder Strict { get; }
        IStringConstraintBuilder StringCompare(StringCompareMode CompareMethod = StringCompareMode.BinaryCompare);

        new IConstraintBuilder EqualTo(object Expected);
        new IConstraintBuilder LessThan(object Expected);
        new IConstraintBuilder LessThanOrEqualTo(object Expected);
        new IConstraintBuilder GreaterThan(object Expected);
        new IConstraintBuilder GreaterThanOrEqualTo(object Expected);

        new IConstraintBuilder Null { get; } // -> umleiten zu DBNull
        IConstraintBuilder Nothing { get; } // -> umleiten zu Null
        new IConstraintBuilder Empty { get; }

        new IConstraintBuilder Not { get; }
        IConstraintBuilder IsNot { get; }
    }

    [ComVisible(true)]
    [Guid("16A0BFAE-49E8-42C7-8AD0-0A340F53264C")]
    public interface IStringConstraintBuilder : IConstraintBuilder, AccUnit.Assertions.IStringConstraintBuilder
    {
        new IConstraintBuilder EqualTo(string Expected);
        new IConstraintBuilder LessThan(string Expected);
        new IConstraintBuilder LessThanOrEqualTo(string Expected);
        new IConstraintBuilder GreaterThan(string Expected);
        new IConstraintBuilder GreaterThanOrEqualTo(string Expected);

        new IConstraintBuilder Empty { get; }

        new IStringConstraintBuilder Not { get; }
        new IStringConstraintBuilder IsNot { get; }
    }

    [ComVisible(true)]
    [Guid("19CFF6F1-9195-4FFE-A685-F26957915EC9")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".ConstraintBuilder")]
    public class ConstraintBuilder : AccUnit.Assertions.ConstraintBuilder, IConstraintBuilder
    {
        public ConstraintBuilder() : base() { }
        public ConstraintBuilder(bool strict) : base(strict) { }

        public IConstraintBuilder Strict { get { return new ConstraintBuilder(true); } }
        public IStringConstraintBuilder StringCompare(StringCompareMode CompareMethod = StringCompareMode.BinaryCompare) {

            StringComparison stringComparison = StringComparison.InvariantCulture;

            switch (CompareMethod)
            {
                case StringCompareMode.BinaryCompare:
                    stringComparison = StringComparison.InvariantCulture;
                    break;
                case StringCompareMode.TextCompare:
                    stringComparison = StringComparison.InvariantCultureIgnoreCase;
                    break;
            }

            return new StringConstraintBuilder(stringComparison); 
        }

        public new IConstraintBuilder EqualTo(object expected)
        {
            return (IConstraintBuilder)base.EqualTo(expected);
        }

        public new IConstraintBuilder GreaterThan(object expected)
        {
            return (IConstraintBuilder)base.GreaterThan(expected);
        }

        public new IConstraintBuilder GreaterThanOrEqualTo(object expected)
        {
            return (IConstraintBuilder)base.GreaterThanOrEqualTo(expected);
        }

        public new IConstraintBuilder LessThan(object expected)
        {
            return (IConstraintBuilder)base.LessThan(expected);
        }

        public new IConstraintBuilder LessThanOrEqualTo(object expected)
        {
            return (IConstraintBuilder)base.LessThanOrEqualTo(expected);
        }

        new public IConstraintBuilder Null
        {
            get
            {
                return (IConstraintBuilder)base.DBNull;
            }
        }

        public IConstraintBuilder Nothing
        {
            get
            {
                return (IConstraintBuilder)base.Null;
            }
        }

        new public IConstraintBuilder Empty
        {
            get
            {
                return (IConstraintBuilder)base.Empty;
            }
        }

        public new IConstraintBuilder Not
        {
            get
            {
                return (IConstraintBuilder)base.Not;
            }
        }

        public IConstraintBuilder IsNot
        {
            get
            {
                return (IConstraintBuilder)base.Not;
            }
        }
    }

    [ComVisible(true)]
    [Guid("35D18449-6FDE-479D-B2C5-BE1BFE7978AE")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".ConstraintBuilder")]
    public class StringConstraintBuilder : AccUnit.Assertions.StringConstraintBuilder, IStringConstraintBuilder
    {
        public StringConstraintBuilder(StringComparison CompareMethod) : base(CompareMethod) { }

        public new IConstraintBuilder EqualTo(string expected)
        {
            return (IConstraintBuilder)base.EqualTo(expected);
        }

        public new IConstraintBuilder GreaterThan(string expected)
        {
            return (IConstraintBuilder)base.GreaterThan(expected);
        }

        public new IConstraintBuilder GreaterThanOrEqualTo(string expected)
        {
            return (IConstraintBuilder)base.GreaterThanOrEqualTo(expected);
        }

        public new IConstraintBuilder LessThan(string expected)
        {
            return (IConstraintBuilder)base.LessThan(expected);
        }

        public new IConstraintBuilder LessThanOrEqualTo(string expected)
        {
            return (IConstraintBuilder)base.LessThanOrEqualTo(expected);
        }

        public IStringConstraintBuilder StringCompare(StringCompareMode CompareMethod = StringCompareMode.BinaryCompare)
        {
            throw new NotImplementedException();
        }

        IConstraintBuilder IConstraintBuilder.EqualTo(object Expected)
        {
            return (IConstraintBuilder)base.EqualTo(Expected.ToString());
        }

        IConstraintBuilder IConstraintBuilder.LessThan(object Expected)
        {
            return (IConstraintBuilder)base.LessThan(Expected.ToString());
        }

        IConstraintBuilder IConstraintBuilder.LessThanOrEqualTo(object Expected)
        {
            return (IConstraintBuilder)base.LessThanOrEqualTo(Expected.ToString());
        }

        IConstraintBuilder IConstraintBuilder.GreaterThan(object Expected)
        {
            return (IConstraintBuilder)base.GreaterThan(Expected.ToString());
        }

        IConstraintBuilder IConstraintBuilder.GreaterThanOrEqualTo(object Expected)
        {
            return (IConstraintBuilder)base.GreaterThanOrEqualTo(Expected.ToString());
        }

        new public IConstraintBuilder Empty
        {
            get
            {
                return (IStringConstraintBuilder)base.Empty;
            }
        }

        public new IStringConstraintBuilder Not
        {
            get
            {
                return (IStringConstraintBuilder)base.Not;
            }
        }

        public IStringConstraintBuilder IsNot
        {
            get
            {
                return (IStringConstraintBuilder)base.Not;
            }
        }

        IConstraintBuilder IConstraintBuilder.Strict {  get { return new ConstraintBuilder(true); } }

        IConstraintBuilder IConstraintBuilder.Null { get { return (IConstraintBuilder)base.Null; } }
        IConstraintBuilder IConstraintBuilder.Nothing { get { return (IConstraintBuilder)base.Null; } }

        IConstraintBuilder IConstraintBuilder.Not { get { return Not; } }

        IConstraintBuilder IConstraintBuilder.IsNot { get { return Not; } }
    }
}
