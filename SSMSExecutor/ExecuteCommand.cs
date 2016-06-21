using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.SqlServer.Management.UI.VSIntegration;

using EnvDTE;
using EnvDTE80;

namespace VSIXProject2
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ExecutorCommand
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("746c2fb4-20a2-4d26-b95d-f8db97c16875");

        private readonly Package package;

        private ExecutorCommand(Package package)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;

            OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(this.Command_Exec, menuCommandID);

                menuItem.BeforeQueryStatus += Command_QueryStatus;

                commandService.AddCommand(menuItem);
            }
        }

        public static ExecutorCommand Instance
        {
            get;
            private set;
        }

        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        public static void Initialize(Package package)
        {
            Instance = new ExecutorCommand(package);
        }

        private void Command_QueryStatus(object sender, EventArgs e)
        {
            var menuCommand = sender as OleMenuCommand;

            if (menuCommand != null)
            {
                menuCommand.Enabled = false;
                menuCommand.Supported = false;

                if (ServiceCache.ExtensibilityModel.ActiveDocument != null)
                {
                    var doc = (ServiceCache.ExtensibilityModel.ActiveDocument.DTE as DTE2).ActiveDocument;

                    if (doc != null)
                    {
                        menuCommand.Enabled = true;
                        menuCommand.Supported = true;
                    }
                }
            }
        }

        private void Command_Exec(object sender, EventArgs e)
        {
            EnvDTE.Document document = Helpers.TextDocument();

            if (document != null)
            {
                var executor = new Executor(document);

                executor.ExecuteCurrentStatement();
            }
        }
    }
}
