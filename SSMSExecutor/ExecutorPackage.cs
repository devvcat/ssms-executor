using System;
using System.ComponentModel.Design;
using System.Diagnostics;
//using System.Diagnostics.CodeAnalysis;
//using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
//using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
//using Microsoft.VisualStudio.Shell.Interop;
//using Microsoft.Win32;
//using Microsoft.SqlServer.Management;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.Win32;
using EnvDTE80;
using EnvDTE;

namespace Devvcat.SSMS
{
    [Guid(ExecutorPackage.PackageGuidString)]
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public class ExecutorPackage : Package
    {
        public const string PackageGuidString = "a64d9865-b938-4543-bf8f-a553cc4f67f3";

        static ExecutorPackage()
        {
            var path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SSMSExecutor");

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            Trace.Listeners.Clear();
            Trace.Listeners.Add(new TextWriterTraceListener(Path.Combine(path, "log.txt")));

#if DEBUG
            Trace.Listeners.Add(new ConsoleTraceListener());
#endif

            Trace.AutoFlush = true;
        }

        protected override void Initialize()
        {
            base.Initialize();

            Trace.TraceInformation("Call 'ExecutorCommand.Initialize(...)'");

            var service = (DTE2)GetService(typeof(DTE));
            ExecutorCommand.Initialize(this, service);
        }

        protected override int QueryClose(out bool canClose)
        {
            SetSkipLoading();

            return base.QueryClose(out canClose);
        }

        void SetSkipLoading()
        {
            try
            {
                var location = GetType().Assembly.Location;

                var m = Regex.Match(location, @"\\(\d{3})\\");
                if (m.Success)
                {
                    var version = int.Parse(m.Groups[1].Value) / 10;

                    SetSkipKey(version, PackageGuidString);
                }
                else
                {
                    SetSkipKey(12, PackageGuidString);
                    SetSkipKey(13, PackageGuidString);
                    SetSkipKey(14, PackageGuidString);
                }
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.Message);
                if (ex.InnerException != null) Trace.TraceError(ex.InnerException.Message);
            }
        }

        void SetSkipKey(float version, string packageGuidString)
        {
            var key = "Software\\Microsoft\\SQL Server Management Studio\\{0:0.0}\\Packages\\{{{1}}}";
            var registryKey = Registry.CurrentUser.CreateSubKey(string.Format(key, version, packageGuidString));

            registryKey.SetValue("SkipLoading", 1, RegistryValueKind.DWord);
            registryKey.Close();
        }
    }
}
