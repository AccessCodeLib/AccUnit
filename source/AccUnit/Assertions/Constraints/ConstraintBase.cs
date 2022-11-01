namespace AccessCodeLib.AccUnit.Assertions.Constraints
{
    abstract class ConstraintBase : IConstraint
    {
        protected IConstraint FirstChild { get; set; }
        protected string CompareText { get; set; }

        protected IConstraint _child;
        public IConstraint Child { get
            {
                return Child;
            }
            set {
                if (FirstChild == null)
                {
                    FirstChild = value;
                    _child = value;
                }
                else
                {
                    _child.Child = value;
                    _child = value;
                }
            }
        }

        public IMatchResult Matches(object actual)
        {
            if (FirstChild != null)
            {
                return Compare(FirstChild.Matches(actual));
            }

            return Compare(actual);
        }

        protected abstract IMatchResult Compare(object actual);
    }
}
