using System;
using System.ComponentModel.Design;
using System.Diagnostics;
//using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using Microsoft.SqlServer.Management;

namespace Devvcat.SSMS
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideAutoLoad(VSConstants.UICONTEXT.EmptySolution_string)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(ExecutorPackage.PackageGuidString)]
    //[SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    public sealed class ExecutorPackage : SqlPackage
    {
        public const string PackageGuidString = "a64d9865-b938-4543-bf8f-a553cc4f67f3";

        public ExecutorPackage()
        {
        }

        private void SetSkipLoading()
        {
            try
            {
                RegistryKey registryKey = Registry.CurrentUser.CreateSubKey(string.Format("Software\\Microsoft\\SQL Server Management Studio\\13.0\\Packages\\{{{0}}}", PackageGuidString));
                registryKey.SetValue("SkipLoading", 1, RegistryValueKind.DWord);
                registryKey.Close();
            }
            catch
            { }
        }

        #region Package Members

        protected override void Initialize()
        {
            ExecutorCommand.Initialize(this);
            base.Initialize();
        }

        protected override int QueryClose(out bool canClose)
        {
            var token = base.QueryClose(out canClose);
            SetSkipLoading();
            return token;
        }

        #endregion
    }
}
