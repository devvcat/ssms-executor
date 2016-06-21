using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.SqlServer.Management.UI.VSIntegration;

namespace SSMSExecutor
{
    public static class Helper
    {
        public static EnvDTE.TextDocument TextDocument()
        {
            EnvDTE.TextDocument doc = null;

            if (ServiceCache.ExtensibilityModel.ActiveDocument != null)
            {
                doc = (EnvDTE.TextDocument)ServiceCache.ExtensibilityModel.ActiveDocument.Object(null);
            }

            return doc;
        }
    }
}
