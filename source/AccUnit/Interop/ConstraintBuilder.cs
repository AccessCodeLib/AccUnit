using System.Runtime.InteropServices;
using System;

namespace AccessCodeLib.AccUnit.Interop
{
    [ComVisible(true)]
    [Guid("0011016E-BF37-4CB7-9A62-58DD78292550")]
    public interface IConstraintBuilder : AccUnit.Assertions.IConstraintBuilder
    {
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
    [Guid("19CFF6F1-9195-4FFE-A685-F26957915EC9")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId(Constants.ProgIdLibName + ".IsConstraints")]
    public class ConstraintBuilder : AccUnit.Assertions.ConstraintBuilder, IConstraintBuilder
    {
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
            get {
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
}
