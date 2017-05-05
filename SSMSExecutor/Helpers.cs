using Microsoft.SqlServer.Management.UI.VSIntegration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Devvcat.SSMS
{
    public static class Helpers
    {
        public static EnvDTE.Document TextDocument()
        {
            EnvDTE.Document document = null;

            if (ServiceCache.ExtensibilityModel.ActiveDocument != null)
            {
                document = ServiceCache.ExtensibilityModel.ActiveDocument.DTE.ActiveDocument; // as EnvDTE.TextDocument);
            }

            return document;
        }
    }
}
