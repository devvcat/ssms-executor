using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.IO;
using System.Xml.Linq;

namespace SSMSExecutor
{
    [RunInstaller(true)]
    public partial class InstallHelper : System.Configuration.Install.Installer
    {
        public InstallHelper()
        {
            InitializeComponent();
        }

        protected override void OnAfterInstall(IDictionary savedState)
        {
            var fileInfo = new FileInfo
                (System.Reflection.Assembly.GetExecutingAssembly().Location);

            var appDataFolder = Context.Parameters["AppData"];
            var addIn = Path.Combine(appDataFolder, @"Microsoft\MSEnvShared\AddIns", Path.ChangeExtension(fileInfo.Name, ".AddIn"));

            if (File.Exists(addIn))
            {
                var xdoc = XDocument.Load(addIn);
                XNamespace ns = "http://schemas.microsoft.com/AutomationExtensibility";

                var assemblyElement = xdoc.Descendants(ns + "Assembly").FirstOrDefault();

                if (assemblyElement != null)
                {
                    assemblyElement.Value = fileInfo.FullName;
                    xdoc.Save(addIn);
                }
            }
            else
            {
#if DEBUG
                System.Diagnostics.Debugger.Launch();
#endif
            }

            base.OnAfterInstall(savedState);
        }
    }
}
