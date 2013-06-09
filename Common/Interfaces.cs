using System;
using System.Data;
using System.Configuration;
using System.Reflection;
using EnvDTE;
using System.Diagnostics.CodeAnalysis;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {

    /// <summary>Interface describing extract to resource actions.</summary>
    /// <remarks>These actions are responsible for filtering resource files supported under current project and document and creating references to resource entries.</remarks>
    public interface IExtractResourceAction {

        /// <summary>Gets priority of the action.</summary>
        /// <remarks>If there are multiple actions supporting the same item, action with the highest priority will be selected</remarks>
        int Priority { get; }

        /// <summary>Gets the default relative path for resource file.</summary>
        string DefaultResourceFilePath { get; }

        /// <summary>Queries if this action supports the provided project item and its containing project</summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns>True if the action suppoors provided project item in context of its parent project; false otherwise</returns>
        bool QuerySupportForProject(EnvDTE.ProjectItem item);

        /// <summary>This method will be used for filtering resource files displayed to user</summary>
        /// <param name="item">Project item for the resource file</param>
        /// <returns>true if resource file is valid and should be displayed to user</returns>
        bool IsValidResourceFile(EnvDTE.ProjectItem item);

        /// <summary>This method should update properties on a recently created resource file so that it is correctly supported by the same instance of <see cref="IExtractResourceAction"/></summary>
        /// <param name="item">Project item for the resource file</param>
        void UpdateResourceFileProperties(EnvDTE.ProjectItem item);

        /// <summary>Returns the code reference to resource specified in the parameters</summary>
        /// <param name="file">Resource file containing the resource</param>
        /// <param name="resourceName">Name of the resource</param>
        /// <param name="project">Project current file belongs to</param>
        /// <param name="string">String being extracted</param>
        /// <returns>a piece of code that would reference to the resource provided</returns>
        /// <remarks>This method does not verify if resource actually exists</remarks>
        string GetResourceReference(ResourceFile file, string resourceName, EnvDTE.Project project, BaseHardCodedString @string);
    }
}