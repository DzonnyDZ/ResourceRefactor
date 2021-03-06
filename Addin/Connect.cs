/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Reflection;
using System.Resources;
using System.Windows.Forms;
using EnvDTE;
using EnvDTE80;
using Extensibility;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VSPowerToys.ResourceRefactor.Common;

namespace Microsoft.VSPowerToys.ResourceRefactor {
    /// <summary>The object for implementing an Add-in.</summary>
    /// <seealso class='IDTExtensibility2' />
    [System.Runtime.InteropServices.ComVisible(true)]
    public class Connect : IDTExtensibility2, IDTCommandTarget {
        private DTE2 applicationObject;
        private AddIn addInInstance;

        /// <summary>Implements the <see cref="IDTExtensibility2.OnConnection">OnConnection</see> method of the <see cref="IDTExtensibility2"/> interface. Receives notification that the Add-in is being loaded.</summary>
        /// <param name='application'>Root object of the host application.</param>
        /// <param name='connectMode'>Describes how the Add-in is being loaded.</param>
        /// <param name='addInInst'>Object representing this Add-in.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom) {
            if (Application == null) {
                throw new ArgumentNullException("Application");
            }
            if (AddInInst == null) {
                throw new ArgumentNullException("AddInInst");
            }
            applicationObject = (DTE2)Application;
            addInInstance = (AddIn)AddInInst;

            if (ConnectMode == ext_ConnectMode.ext_cm_UISetup) {
                this.SetupRefactorActions();
            }

        }

        /// <summary>Implements the <see cref="IDTExtensibility2.OnDisconnection">OnDisconnection</see> method of the <see cref="IDTExtensibility2"/> interface. Receives notification that the Add-in is being unloaded.</summary>
        /// <param name='disconnectMode'>Describes how the Add-in is being unloaded.</param>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnDisconnection(ext_DisconnectMode disconnectMode, ref Array custom) {
        }

        /// <summary>Implements the <see cref="IDTExtensibility2.OnAddInsUpdate">OnAddInsUpdate</see> method of the <see cref="IDTExtensibility2"/> interface. Receives notification when the collection of Add-ins has changed.</summary>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnAddInsUpdate(ref Array custom) {
        }

        /// <summary>Implements the <see cref="IDTExtensibility2.OnStartupComplete">OnStartupComplete</see> method of the <see cref="IDTExtensibility2"/> interface. Receives notification that the host application has completed loading.</summary>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnStartupComplete(ref Array custom) {
        }

        /// <summary>Implements the <see cref="IDTExtensibility2.OnBeginShutdown">OnBeginShutdown</see> method of the <see cref="IDTExtensibility2"/> interface. Receives notification that the host application is being unloaded.</summary>
        /// <param name='custom'>Array of parameters that are host application specific.</param>
        /// <seealso class='IDTExtensibility2' />
        public void OnBeginShutdown(ref Array custom) {
        }

        #region IDTCommandTarget Members

        /// <summary>Handles the command execution.</summary>
        /// <param name="cmdName">The name of the command to execute.</param>
        /// <param name="executeOption">A <see cref="vsCommandExecOption"/> constant specifying the execution options.</param>
        /// <param name="variantIn">A value passed to the command.</param>
        /// <param name="variantOut">A value passed back to the invoker <see cref="Exec"/> method after the command executes.</param>
        /// <param name="handled">Returns value indicating if the command has been handle</param>
        /// <remarks>First method checks if user has selected a part of string literal in the code, if so method invokes RefactorStringDialog.</remarks>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        public void Exec(string cmdName, vsCommandExecOption executeOption, ref object variantIn, ref object variantOut, ref bool handled) {
            if (cmdName == null) {
                throw new ArgumentNullException("CmdName");
            }
            if (cmdName == typeof(Connect).FullName + "." + Strings.RefactorCommandName ) {
                TextSelection selection = (TextSelection)(applicationObject.ActiveDocument.Selection);
                if (applicationObject.ActiveDocument.ProjectItem.Object != null)
                {

                    Common.BaseHardCodedString stringInstance = BaseHardCodedString.GetHardCodedString(applicationObject.ActiveDocument);

                    if (stringInstance == null)
                    {
                        MessageBox.Show(
                            Strings.UnsupportedFile + " (" + applicationObject.ActiveDocument.Language + ")",
                            Strings.WarningTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return;
                    }

                    Common.MatchResult scanResult = stringInstance.CheckForHardCodedString(
                       selection.Parent,
                       selection.AnchorPoint.AbsoluteCharOffset - 1,
                       selection.BottomPoint.AbsoluteCharOffset - 1);

                    if (!scanResult.Result && selection.AnchorPoint.AbsoluteCharOffset < selection.BottomPoint.AbsoluteCharOffset) {
                        scanResult.StartIndex = selection.AnchorPoint.AbsoluteCharOffset - 1;
                        scanResult.EndIndex = selection.BottomPoint.AbsoluteCharOffset - 1;
                        scanResult.Result = true;
                    }
                    if (scanResult.Result) {
                        stringInstance = stringInstance.CreateInstance(applicationObject.ActiveDocument.ProjectItem, scanResult.StartIndex, scanResult.EndIndex);
                        if (stringInstance != null && stringInstance.Parent != null) {
                            PerformAction(stringInstance);
                        }
                    } else {
                        MessageBox.Show(Strings.NotStringLiteral, Strings.WarningTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>Shows "Extract to Resource" dialog to user and extracts the string to resource files</summary>
        /// <param name="stringInstance">String to refactor</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        private void PerformAction(Common.BaseHardCodedString stringInstance) {
            ExtractToResourceActionSite site = new ExtractToResourceActionSite(stringInstance);
            if (site.ActionObject != null) {
                RefactorStringDialog dialog = new RefactorStringDialog();
                dialog.ShowDialog(site);
            } else {
                MessageBox.Show(Strings.UnsupportedFile, Strings.WarningTitle, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>Returns the current status (enabled, disabled, hidden, and so forth) of the specified named command.</summary>
        /// <param name="cmdName">The name of the command to check.</param>
        /// <param name="neededText">A <see cref="vsCommandStatusTextWanted"/> constant specifying if information is returned from the check, and if so, what type of information is returned.</param>
        /// <param name="statusOption">A <see cref="vsCommandStatus"/> specifying the current status of the command.</param>
        /// <param name="commandText">The text to return if <see cref="vsCommandStatusTextWantedStatus"/> is specified.</param>
        public void QueryStatus(string cmdName, vsCommandStatusTextWanted neededText, ref vsCommandStatus statusOption, ref object commandText) {
            if (cmdName == null) {
                throw new ArgumentNullException("CmdName");
            }
            if (neededText == vsCommandStatusTextWanted.vsCommandStatusTextWantedNone) {
                if (cmdName == typeof(Connect).FullName +"." + Strings.RefactorCommandName) {
                    statusOption = vsCommandStatus.vsCommandStatusUnsupported | vsCommandStatus.vsCommandStatusInvisible;
                    // Check if current document is contained in a project
                    if (applicationObject.ActiveDocument != null &&
                        applicationObject.ActiveDocument.ProjectItem != null &&
                        applicationObject.ActiveDocument.ProjectItem.ContainingProject != null &&
                        !applicationObject.ActiveDocument.ProjectItem.ContainingProject.Kind.Equals("{66A2671D-8FB5-11D2-AA7E-00C04F688DDE}"/*EnvDTE.Constants.vsProjectKindMisc*/)) {
                        // Check if current window is not a designer
                        if (!(applicationObject.ActiveWindow.Object is System.ComponentModel.Design.IDesignerHost)) {
                            switch (applicationObject.ActiveDocument.Language) {
                                case "CSharp":
                                case "Basic":
                                case "XAML":
                                    statusOption = vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                                    break;
                                case "HTML":
                                    if (applicationObject.ActiveDocument.Name.EndsWith(".cshtml", StringComparison.CurrentCultureIgnoreCase) ||
                                        applicationObject.ActiveDocument.Name.EndsWith(".vbhtml", StringComparison.CurrentCultureIgnoreCase) ||
                                        applicationObject.ActiveDocument.Name.EndsWith(".aspx", StringComparison.CurrentCultureIgnoreCase) || applicationObject.ActiveDocument.Name.EndsWith(".ascx", StringComparison.CurrentCultureIgnoreCase) || applicationObject.ActiveDocument.Name.EndsWith(".master", StringComparison.CurrentCultureIgnoreCase)) {
                                        statusOption = vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
        }

        #endregion

        /// <summary>Finds the localized menu name from English menu name</summary>
        /// <param name="name">Name of the menu</param>
        /// <returns>Localized name</returns>
        /// <remarks>We have to catch all exceptions since GetString method can throw a very extensive list of exceptions</remarks>
        [SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes")]
        private string FindLocalizedMenuName(string name) {
            string menuName;
            ResourceManager resourceManager = new ResourceManager("Microsoft.VSPowerToys.ResourceRefactor.CommandBar", Assembly.GetExecutingAssembly());
            CultureInfo cultureInfo = new System.Globalization.CultureInfo(this.applicationObject.LocaleID);
            string resourceName = String.Concat(cultureInfo.TwoLetterISOLanguageName, name);
            try {
                menuName = resourceManager.GetString(resourceName);
            } catch {
                //We tried to find a localized version of the word, but one was not found.
                //Default to the en-US word, which may work for the current culture.
                menuName = name;
            }
            return menuName;
        }

        /// <summary>Registers refactoring actions with Visual Studio</summary>
        private void SetupRefactorActions() {
            object[] contextGUIDS = new object[] { };
            Commands2 commands = (Commands2)applicationObject.Commands;
            Command refactorCommand = null;

            IList<CommandBar> targets = new List<CommandBar>();
            try {
                try {
                    CommandBar menuBarCommandBar = ((CommandBars)applicationObject.CommandBars)["MenuBar"];
                    targets.Add(((CommandBarPopup)menuBarCommandBar.Controls[FindLocalizedMenuName("Refactor")]).CommandBar);
                } catch { }
                try {
                    targets.Add(((CommandBars)applicationObject.CommandBars)["Refactor"]);
                } catch { }
                try {
                    targets.Add(((CommandBars)applicationObject.CommandBars)["XAML Editor"]);
                } catch { }
                try {
                    targets.Add(((CommandBars)applicationObject.CommandBars)["ASPX Context"]);
                } catch { }
                try {
                    targets.Add(((CommandBars)applicationObject.CommandBars)["ASPX Code Context"]);
                } catch { }
                try {
                    targets.Add(((CommandBars)applicationObject.CommandBars)["ASPX VB Code Context"]);
                } catch { }
                try {
                    targets.Add(((CommandBars)applicationObject.CommandBars)["HTML Context"]);
                } catch { }

                foreach (var target in targets) {
                    while (true) {
                        try {
                            CommandBarControl ctl = (CommandBarControl)(target.Controls[Strings.RefactorCommandText]);
                            ctl.Delete(false);
                        } catch (ArgumentException) {
                            break;
                        }
                    }
                }

                // Create refactor command if necessary
                try {
                    refactorCommand = commands.AddNamedCommand2(addInInstance,
                        Strings.RefactorCommandName,
                        Strings.RefactorCommandText,
                        Strings.RefactorCommandToolTip,
                        true, 69, ref contextGUIDS, (int)(vsCommandStatus.vsCommandStatusEnabled | vsCommandStatus.vsCommandStatusSupported), (int)vsCommandStyle.vsCommandStyleText, vsCommandControlType.vsCommandControlTypeButton);
                    if (!String.IsNullOrEmpty(Strings.RefactorCommandHotkey)) {
                        refactorCommand.Bindings = new object[] { Strings.RefactorCommandHotkey };
                    }
                } catch (ArgumentException) {
                    refactorCommand = commands.Item(addInInstance.ProgID + "." + Strings.RefactorCommandName, -1);
                }

                foreach (var target in targets) {
                    refactorCommand.AddControl(target, target.Controls.Count + 1);
                }

            } catch (Exception e) {
                System.Diagnostics.Trace.TraceError(e.ToString());
                throw;
            }
        }
    }
}