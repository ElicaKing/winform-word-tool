using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordOpenXmlClassLibrary
{
    public class GenerateBookmarkEnd
    {
        // Creates an BookmarkEnd instance and adds its children.
        public BookmarkEnd Create()
        {
            BookmarkEnd bookmarkEnd = new BookmarkEnd() { Id = "0" };
            return bookmarkEnd;
        }

    }
}
