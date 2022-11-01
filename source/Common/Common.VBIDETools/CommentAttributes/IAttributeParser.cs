namespace AccessCodeLib.Common.VBIDETools.CommentAttributes
{
    public interface IAttributeParser<out TAttribute> where TAttribute : CommentAttribute
    {
        TAttribute Parse(string comment);
    }
}