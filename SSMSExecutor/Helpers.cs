using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE80;

namespace Devvcat.SSMS
{
    static class Helpers
    {
        public static bool HasActiveDocument(this DTE2 dte)
        {
            if (dte.ActiveDocument != null)
            {
                var doc = (dte.ActiveDocument.DTE)?.ActiveDocument;
                return doc != null;
            }

            return false;
        }

        public static EnvDTE.Document GetDocument(this DTE2 dte)
        {
            if (dte.HasActiveDocument())
            {
                return (dte.ActiveDocument.DTE)?.ActiveDocument;
            }

            return null;
        }
    }
}
