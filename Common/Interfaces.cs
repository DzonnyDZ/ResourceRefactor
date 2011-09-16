using System;
using System.Data;
using System.Configuration;
using System.Reflection;
using EnvDTE;
using System.Diagnostics.CodeAnalysis;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common
{

    /// <summary>
    /// Interface describing a string refactoring option that can be selected by user.
    /// </summary>
    /// <remarks>This interface describes all necessary information to create a menu item and invoke refactor string command</remarks>
    public interface IStringRefactorOption
    {
        /// <summary>
        /// Command name to be used in Visual Studio. Actual command name will have ProgID prefixed.
        /// </summary>
        /// <remarks>If an item with the same commandname is already registered, the second option will be 
        /// ignored.</remarks>
        string CommandName { get; }

        /// <summary>
        /// Description of the refactoring option, this will displayed as a tool tip in the menu entry
        /// </summary>
        string Description { get; }

        /// <summary>
        /// Text of the menu entry which will be shown in "Refactor" submenus
        /// </summary>
        string MenuText { get; }

        /// <summary>
        /// Default hotkey for the refactor action. This should be in the format described at
        /// http://msdn2.microsoft.com/en-us/library/envdte.command.bindings(VS.80).aspx
        /// 
        /// If left empty or null, no hotkey will be registered.
        /// </summary>
        string Hotkey { get; }

        /// <summary>
        /// Queries refactoring option to check if it supports the provided string instance
        /// </summary>
        /// <param name="stringInstance">String instance to refactor</param>
        /// <returns>true if string instance can be refactored using this option, false otherwise</returns>
        bool QuerySupportForString(BaseHardCodedString stringInstance);

        /// <summary>
        /// Invokes refactoring option to perform necessary steps for refactoring
        /// </summary>
        /// <param name="stringInstance">String instance to refactor</param>
        void PerformAction(BaseHardCodedString stringInstance);

        /// <summary>
        /// Wrapper for IDTCommandTarget.QueryStatus since each option has to determine their status individually.
        /// </summary>
        /// <param name="commandName"></param>
        /// <param name="NeededText"></param>
        /// <param name="StatusOption"></param>
        /// <param name="CommandText"></param>
        /// <param name="applicationObject"></param>
        /// <remarks>Design messages are ignored since this method is coming from IDTCommandTarget interface</remarks>
        [SuppressMessage("Microsoft.Design", "CA1007:UseGenericsWhereAppropriate")]
        [SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "2#")]
        [SuppressMessage("Microsoft.Design", "CA1045:DoNotPassTypesByReference", MessageId = "3#")]
        void QueryStatus(string commandName, EnvDTE.vsCommandStatusTextWanted neededText, ref EnvDTE.vsCommandStatus statusOption, ref object commandText, EnvDTE80.DTE2 applicationObject);
        
    }

    /// <summary>
    /// Interface describing extract to resource actions. These actions are responsible for filtering resource files
    /// supported under current project and document and creating references to resource entries.
    /// </summary>
    public interface IExtractResourceAction
    {

        /// <summary>
        /// Priority of the action. If there are multiple actions supporting the same item, action with
        /// the highest priority will be selected
        /// </summary>
        int Priority
        {
            get;
        }

        /// <summary>
        /// Gets the default relative path for resource file.
        /// </summary>
        string DefaultResourceFilePath
        {
            get;
        }

        /// <summary>
        /// Queries if this action supports the provided project item and its containing project
        /// </summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns></returns>
        bool QuerySupportForProject(EnvDTE.ProjectItem item);

        /// <summary>
        /// This method will be used for filtering resource files displayed to user
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        /// <returns>true if resource file is valid and should be displayed to user</returns>
        bool IsValidResourceFile(EnvDTE.ProjectItem item);

        /// <summary>
        /// This method should update properties on a recently created resource file so that it
        /// is correctly supported by the same instance of IExtractResourceAction
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        void UpdateResourceFileProperties(EnvDTE.ProjectItem item);

        /// <summary>
        /// Returns the code reference to resource specified in the parameters
        /// </summary>
        /// <param name="file">Resource file containing the resource</param>
        /// <param name="resourceName">Name of the resource</param>
        /// <param name="project">Project current file belongs to</param>
        /// <returns>a piece of code that would reference to the resource provided</returns>
        /// <remarks>This method does not verify if resource actually exists</remarks>
        string GetResourceReference(ResourceFile file, string resourceName, EnvDTE.Project project);
    }

    

}
