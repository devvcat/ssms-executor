using System;
using System.ComponentModel.Design;

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
        public const int ExecuteStatementCommandId = 0x0100;
        public const int ExecuteInnerStatementCommandId = 0x0101;

        public static readonly Guid CommandSet = new Guid("746c2fb4-20a2-4d26-b95d-f8db97c16875");

        readonly Package package;
        readonly DTE2 dte;

        private ExecutorCommand(Package package)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            this.dte = (DTE2)ServiceProvider.GetService(typeof(DTE));

            if (ServiceProvider.GetService(typeof(IMenuCommandService)) is OleMenuCommandService commandService)
            {
                CommandID menuCommandID;
                OleMenuCommand menuCommand;

                // Create execute current statement menu item
                menuCommandID = new CommandID(CommandSet, ExecuteStatementCommandId);
                menuCommand = new OleMenuCommand(Command_Exec, menuCommandID);
                menuCommand.BeforeQueryStatus += Command_QueryStatus;
                commandService.AddCommand(menuCommand);

                // Create execute inner satetement menu item
                menuCommandID = new CommandID(CommandSet, ExecuteInnerStatementCommandId);
                menuCommand = new OleMenuCommand(Command_Exec, menuCommandID);
                menuCommand.BeforeQueryStatus += Command_QueryStatus;
                commandService.AddCommand(menuCommand);
            }
        }

        private IServiceProvider ServiceProvider => package;

        public static ExecutorCommand Instance
        {
            get;
            private set;
        }

        public static void Initialize(Package package)
        {
            Instance = new ExecutorCommand(package);
        }

        private Executor.ExecScope GetScope(int commandId)
        {
            var scope = Executor.ExecScope.Block;
            if (commandId == ExecuteInnerStatementCommandId)
            {
                scope = Executor.ExecScope.Inner;
            }
            return scope;
        }

        private void Command_QueryStatus(object sender, EventArgs e)
        {
            if (sender is OleMenuCommand menuCommand)
            {
                menuCommand.Enabled = false;
                menuCommand.Supported = false;

                if (dte.HasActiveDocument())
                {
                    menuCommand.Enabled = dte.ActiveWindow.HWnd == dte.ActiveDocument.ActiveWindow.HWnd;
                    menuCommand.Supported = true;
                }
            }
        }

        private void Command_Exec(object sender, EventArgs e)
        {
            if (sender is OleMenuCommand menuCommand)
            {
                var executor = new Executor(dte);
                var scope = GetScope(menuCommand.CommandID.ID);

                executor.ExecuteStatement(scope);
            }
        }
    }
}
