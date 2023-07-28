using System;
using System.Collections.Generic;

namespace AccessCodeLib.Common.VBIDETools
{
    public class CommentsBlockReader
    {
        public IEnumerable<string> Read(string commentsBlock)
        {
            if (commentsBlock is null)
                return null;

            if (commentsBlock == "")
                return new[] { "" };

            return GetCommentLineContents(commentsBlock);
        }

        private IEnumerable<string> GetCommentLineContents(string commentsBlock)
        {
            var nonEmptyLines = commentsBlock.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in nonEmptyLines)
            {
                var temp = line.Trim();
                if (temp != "")
                {
                    if (!temp.StartsWith("'"))
                    {
                        throw new NonCommentLineException(temp);
                    }
                    temp = temp.Substring(1).TrimStart();
                    yield return temp;
                }
            }
        }
    }
}