using System;
using System.ComponentModel.Design;
using System.Diagnostics;

using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;

namespace Devvcat.SSMS
{
    /// <summary>
    /// Command handler
    /// </summary>
    sealed class ExecutorCommand
    {
        public const int CommandId = 0x0100;
        public static readonly Guid CommandSet = new Guid("746c2fb4-20a2-4d26-b95d-f8db97c16875");

        readonly Package package;
        readonly DTE2 dte;

        private ExecutorCommand(Package package, DTE2 dte)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            this.dte = dte;

            if (ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                Trace.TraceInformation("Create the 'Executor' command.");

                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new OleMenuCommand(Command_Exec, menuCommandID);

                //menuItem.BeforeQueryStatus += Command_QueryStatus;

                commandService.AddCommand(menuItem);
            }
            else
            {
                Trace.TraceWarning("Failed to find 'MenuCommandService'.");
            }
        }

        private IServiceProvider ServiceProvider => package;

        public static ExecutorCommand Instance
        {
            get;
            private set;
        }

        public static void Initialize(Package package, DTE2 dte)
        {
            Trace.TraceInformation("Create 'ExecutorCommand' instance");
            Instance = new ExecutorCommand(package, dte);
        }

        bool HasActiveDocument()
        {
            if (dte.ActiveDocument !=null)
            {
                var doc = (dte.ActiveDocument.DTE)?.ActiveDocument;
                return doc != null;
            }

            return false;
        }

        private void Command_QueryStatus(object sender, EventArgs e)
        {
            if (sender is OleMenuCommand menuCommand)
            {
                menuCommand.Enabled = false;
                menuCommand.Supported = false;

                if (dte.HasActiveDocument())
                {
                    menuCommand.Enabled = true;
                    menuCommand.Supported = true;
                }
            }
        }

        private void Command_Exec(object sender, EventArgs e)
        {
            Document document = dte.GetDocument();

            if (document != null)
            {
                var executor = new Executor(document);
                
                executor.ExecuteCurrentStatement();
            }
        }
    }
}
