using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddInVSTO.Extensions
{
    public static class ShapeExtensions
    {
        public static MediaBookmark GetBookmark(this MediaBookmarks bookmarks, string name)
        {
            foreach(MediaBookmark bookmark in bookmarks)
            {
                if (bookmark.Name == name) return bookmark;
            }
            return null;
        }
    }
}
