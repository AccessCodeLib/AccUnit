namespace AccessCodeLib.AccUnit
{
    public interface ITestItemTag
    {
        string Name { get; }

        bool Equals(object obj);
        int GetHashCode();
        string ToString();
    }
}