using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Shell;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Devvcat.SSMS.Options
{
    [Guid("5F238789-E306-48AE-93D6-44FCC2EAAC80")]
    internal class GeneralOptionsPage : DialogPage
    {
        [Category("Execute Options")]
        [DisplayName("Execute inline statements")]
        [Description("Execute inline statements instead of batch")]
        [DefaultValue(false)]
        public bool ExecuteInlineStatements { get; set; }

        protected override void OnActivate(CancelEventArgs e)
        {
            ExecuteInlineStatements = Properties.Settings.Default.ExecuteInlineStatements;

            base.OnActivate(e);
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            Properties.Settings.Default.ExecuteInlineStatements = ExecuteInlineStatements;
            Properties.Settings.Default.Save();

            base.OnApply(e);
        }
    }
}
