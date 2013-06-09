using System;
using IO = System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>Basic abstract implementation of <see cref="IExtractResourceAction "/></summary>
    public abstract class ExtractResourceActionBase : IExtractResourceAction {
        /// <summary>Gets priority of the action.</summary>
        public virtual int Priority { get { return 10; } }

        /// <summary>Gets the default relative path for resource file.</summary>
        public virtual string DefaultResourceFilePath { get { return String.Empty; } }

        /// <summary>Queries if this action supports the provided project item and its containing project</summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns>True if the action is supported, false otherwise.</returns>
        public abstract bool QuerySupportForProject(EnvDTE.ProjectItem item);

        /// <summary>This method will be used for filtering resource files displayed to user</summary>
        /// <param name="item">Project item for the resource file</param>
        /// <returns>true if resource file is valid and should be displayed to user</returns>
        public virtual bool IsValidResourceFile(EnvDTE.ProjectItem item) {
            if (item == null) return false;
            try {
                if (item.Properties.Item("Extension").Value.ToString().Equals(".resx", StringComparison.InvariantCultureIgnoreCase)) {

                    if (item.Properties.Item("ItemType").Value.Equals("EmbeddedResource")) {
                        return true;
                    }
                }
                return false;
            } catch (ArgumentException) {
                // This can happen if item does not contain the properties we are looking for.
                return false;
            }
        }

        /// <summary>When overriden in derived class returns the code reference to resource specified in the parameters</summary>
        /// <param name="file">Resource file containing the resource</param>
        /// <param name="resourceName">Name of the resource</param>
        /// <param name="project">Project current file belongs to</param>
        /// <param name="string">String being extracted</param>
        /// <returns>a piece of code that would reference to the resource provided</returns>
        /// <remarks>This method does not verify if resource actually exists</remarks>
        public abstract string GetResourceReference(ResourceFile file, string resourceName, Project project, BaseHardCodedString @string);

        /// <summary>This method should update properties on a recently created resource file so that it is correctly supported by the same instance of <see cref="IExtractResourceAction"/></summary>
        /// <param name="item">Project item for the resource file</param>
        public virtual void UpdateResourceFileProperties(ProjectItem item) {
            if (item != null) {
                item.Properties.Item("CustomTool").Value = "ResXFileCodeGenerator";
                item.Properties.Item("ItemType").Value = "EmbeddedResource";
                //For VB project, we need to set namespace property as well
                item.Properties.Item("CustomToolNamespace").Value = "My.Resources";
            }
        }

        /// <summary>Retruns namespace prefix for a given file</summary>
        /// <param name="file">A file to get namespace prefix for</param>
        /// <returns>Namespace prefix for <paramref name="file"/></returns>
        protected abstract string GetNamespacePrefix(ResourceFile file);

        /// <summary>Checks if provided item is a valid resource file for website projects</summary>
        /// <param name="item">Item to be checked</param>
        /// <returns>true if item is a valid resource file</returns>
        protected static bool CheckResourceFileForWebSites(ProjectItem item) {
            try {
                if (item.Properties.Item("Extension").Value.ToString().Equals(".resx", StringComparison.InvariantCultureIgnoreCase)) {
                    return IO.Path.GetDirectoryName(item.Properties.Item("RelativeURL").Value.ToString()).Equals("App_GlobalResources");
                }
                return false;
            } catch (ArgumentException) {
                // This can happen if item does not contain the properties we are looking for.
                return false;
            }
        }
    }
}
