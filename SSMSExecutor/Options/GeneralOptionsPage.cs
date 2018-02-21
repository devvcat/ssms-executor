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
        [DisplayName("Execute inner statements")]
        [Description("Execute inner statements instead of block")]
        [DefaultValue(false)]
        public bool ExecuteInnerStatements { get; set; }

        protected override void OnActivate(CancelEventArgs e)
        {
            ExecuteInnerStatements = Properties.Settings.Default.ExecuteInnerStatements;

            base.OnActivate(e);
        }

        protected override void OnApply(PageApplyEventArgs e)
        {
            Properties.Settings.Default.ExecuteInnerStatements = ExecuteInnerStatements;
            Properties.Settings.Default.Save();

            base.OnApply(e);
        }
    }
}
