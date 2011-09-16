using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VSPowerToys.ResourceRefactor.Common;
using EnvDTE;
using System.Windows.Forms;
using System.Diagnostics.CodeAnalysis;

namespace Microsoft.VSPowerToys.ResourceRefactor
{
    /// <summary>
    /// This string refactor option extracts to hard coded string to a resource file in the project and replaces
    /// the string with a programmatic reference to the resource entry
    /// </summary>
    public class ExtractToResourceOption : IStringRefactorOption
    {
        #region IStringRefactorAction Members

        /// <summary>
        /// Visual Studio command name for this option
        /// </summary>
        public string CommandName
        {
            get { return Strings.RefactorCommandName; }
        }

        /// <summary>
        /// Description of the option
        /// </summary>
        public string Description
        {
            get { return Strings.RefactorCommandToolTip; }
        }

        /// <summary>
        /// Text to be displayed in the Refactor menu
        /// </summary>
        public string MenuText
        {
            get { return Strings.RefactorCommandText; }
        }

        /// <summary>
        /// Hotkey for the command.
        /// </summary>
        public string Hotkey
        {
            get { return Strings.RefactorCommandHotkey; }
        }
        /// <summary>
        /// Queries the implementation to check if string instance is supported
        /// </summary>
        /// <param name="stringInstance">String instance to check for</param>
        /// <returns>True if string is supported</returns>
        public bool QuerySupportForString(BaseHardCodedString stringInstance)
        {
            return stringInstance is CSharpHardCodedString || stringInstance is VBHardCodedString || stringInstance is XamlHardCodedString;
        }

        /// <summary>
        /// Shows "Extract to Resource" dialog to user and extracts the string to resource files
        /// </summary>
        /// <param name="stringInstance">String to refactor</param>
        [SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions")]
        public void PerformAction(Common.BaseHardCodedString stringInstance)
        {
            ExtractToResourceActionSite site = new ExtractToResourceActionSite(stringInstance);
            if (site.ActionObject != null)
            {
                RefactorStringDialog dialog = new RefactorStringDialog();
                dialog.ShowDialog(site);
            }
            else
            {
                MessageBox.Show(
                                Strings.UnsupportedFile,
                                Strings.WarningTitle,
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Queries the status of the command. This is an implementation of IDTCommandTarget.QueryStatus method
        /// </summary>
        /// <param name="commandName"></param>
        /// <param name="NeededText"></param>
        /// <param name="StatusOption"></param>
        /// <param name="CommandText"></param>
        /// <param name="applicationObject"></param>
        public void QueryStatus(string commandName, EnvDTE.vsCommandStatusTextWanted neededText, ref EnvDTE.vsCommandStatus statusOption, ref object commandText, EnvDTE80.DTE2 applicationObject)
        {
            if (applicationObject == null)
            {
                throw new ArgumentNullException("applicationObject");
            }
            statusOption = vsCommandStatus.vsCommandStatusUnsupported | vsCommandStatus.vsCommandStatusInvisible;
            // Check if current document is contained in a project
            if (applicationObject.ActiveDocument != null &&
                applicationObject.ActiveDocument.ProjectItem != null &&
                applicationObject.ActiveDocument.ProjectItem.ContainingProject != null &&
                !applicationObject.ActiveDocument.ProjectItem.ContainingProject.Kind.Equals(EnvDTE.Constants.vsProjectKindMisc))
            {
                // Check if current window is not a designer
                if (!(applicationObject.ActiveWindow.Object is System.ComponentModel.Design.IDesignerHost))
                {
                    statusOption = vsCommandStatus.vsCommandStatusSupported | vsCommandStatus.vsCommandStatusEnabled;
                }
            }
        }

        #endregion
    }
}
