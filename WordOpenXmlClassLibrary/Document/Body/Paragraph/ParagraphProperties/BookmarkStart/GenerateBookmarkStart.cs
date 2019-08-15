using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace WordOpenXmlClassLibrary
{
    public class GenerateBookmarkStart
    {
        // Creates an BookmarkStart instance and adds its children.
        public BookmarkStart Create()
        {
            BookmarkStart bookmarkStart = new BookmarkStart() { Name = "_GoBack", Id = "0" };
            return bookmarkStart;
        }

    }
}
