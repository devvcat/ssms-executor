using System;
using System.IO;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;

namespace Devvcat.SSMS
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "2.0.1", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class ExecutorPackage : Package
    {
        private const string PackageGuidString = "a64d9865-b938-4543-bf8f-a553cc4f67f3";

        public ExecutorPackage()
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
            Trace.TraceInformation("Entering constructor for {0}", ToString());
        }

        protected override void Initialize()
        {
            base.Initialize();

            Trace.TraceInformation("Call 'ExecutorCommand.Initialize(...)'");

            var service = (EnvDTE80.DTE2)GetService(typeof(EnvDTE.DTE));
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
                var registryKey = UserRegistryRoot.CreateSubKey(
                    string.Format("Packages\\{{{0}}}", PackageGuidString));

                registryKey.SetValue("SkipLoading", 1, RegistryValueKind.DWord);
                registryKey.Close();
            }
            catch (Exception ex)
            {
                Trace.TraceError(ex.Message);
                if (ex.InnerException != null) Trace.TraceError(ex.InnerException.Message);
            }
        }
    }
}
