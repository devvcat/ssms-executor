using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using EnvDTE;

using Extensibility;
using Microsoft.SqlServer.Management.UI.VSIntegration;
using Microsoft.VisualStudio.CommandBars;

namespace SSMSExecutor
{
    public class Connect : IDTExtensibility2, IDTCommandTarget
    {
        private const string COMMAND_NAME = "ExecuteCurrentStatement";
        private const string COMMAND_CAPTION = "Execute Current Statement";

        private EnvDTE.DTE applicationObject;
        private EnvDTE.AddIn addInInstance;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            applicationObject = (EnvDTE.DTE)ServiceCache.ExtensibilityModel;
            addInInstance = (EnvDTE.AddIn)addInInst;

            switch (connectMode)
            {
                case ext_ConnectMode.ext_cm_UISetup:
                    UISetup();
                    break;

                case ext_ConnectMode.ext_cm_Startup:
                    break;

                case ext_ConnectMode.ext_cm_AfterStartup:
                    break;
            }
        }

        public void OnAddInsUpdate( ref Array custom )
        {
        }

        public void OnBeginShutdown( ref Array custom )
        {
        }

        public void OnDisconnection( ext_DisconnectMode RemoveMode, ref Array custom )
        {
        }

        public void OnStartupComplete( ref Array custom )
        {
        }

        private void UISetup()
        {
            Command command = null;

            try
            {
                try
                {
                    command = applicationObject.Commands.Item(
                        string.Format("{0}.{1}", addInInstance.ProgID, "ExecuteCurrentStatement"), -1);

#if DEBUG
                    command.Delete();
                    command = null;
#endif
                }
                catch
                { }

                if (command == null)
                {
                    var contextGuids = new object[0];
                    var commads = applicationObject.Commands;

                    command = commads.AddNamedCommand(
                        addInInstance, COMMAND_NAME, COMMAND_CAPTION, COMMAND_CAPTION, true, 144, null,
                        (int)vsCommandStatus.vsCommandStatusSupported | (int)vsCommandDisabledFlags.vsCommandDisabledFlagsGrey);

                    command.Bindings = "SQL Query Editor::Ctrl+Shift+E";

                    var commandBars = (CommandBars)applicationObject.CommandBars;
                    var commandBarOwner = commandBars["Standard"];

                    var item = (CommandBarButton)command.AddControl(commandBarOwner, commandBarOwner.Controls.Count + 1);
                    item.Caption = COMMAND_CAPTION;
                    item.Style = MsoButtonStyle.msoButtonIcon;
                    item.BeginGroup = true;
                }
            }
            catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
        }

        public void QueryStatus(string commandName, vsCommandStatusTextWanted neededText, ref vsCommandStatus statusOption, ref object commandText)
        {
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone)
            {
                if (commandName == string.Format("{0}.{1}", addInInstance.ProgID, COMMAND_NAME))
                {
                    statusOption = vsCommandStatus.vsCommandStatusSupported;

                    if (ServiceCache.ExtensibilityModel.ActiveDocument != null)
                    {
                        var doc = (EnvDTE.TextDocument)ServiceCache.ExtensibilityModel.ActiveDocument.Object(null);

                        if (doc != null)
                        {
                            statusOption = vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                        }
                    }
                }
                else
                {
                    statusOption = vsCommandStatus.vsCommandStatusUnsupported;
                }
            }
        }

        public void Exec( string commandName, vsCommandExecOption executeOption, ref object variantIn, ref object variantOut, ref bool handled )
        {
            handled = false;

            if (executeOption == vsCommandExecOption.vsCommandExecOptionDoDefault)
            {
                if (commandName == string.Format("{0}.{1}", addInInstance.ProgID, COMMAND_NAME))
                {
                    var doc = Helper.TextDocument();

                    if (doc != null)
                    {
                        var executor = new Executor(doc);

                        executor.ExecuteCurrentStatement();

                        handled = true;
                    }
                }
            }
        }
    }
}
