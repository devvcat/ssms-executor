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
using Devvcat.SSMS.Options;

namespace Devvcat.SSMS
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "2.0.2", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.ShellInitialized_string)]
#if DEBUG
    [ProvideOptionPage(typeof(GeneralOptionsPage), "SSMS Executor", "General", 100, 101, true)]
#endif
    [Guid(PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class ExecutorPackage : Package
    {
        private const string PackageGuidString = "a64d9865-b938-4543-bf8f-a553cc4f67f3";

        public ExecutorPackage()
        {
        }

        protected override void Initialize()
        {
            base.Initialize();

            ExecutorCommand.Initialize(this);
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
            catch
            { }
        }
    }
}
