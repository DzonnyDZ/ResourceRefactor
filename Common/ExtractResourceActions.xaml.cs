using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using EnvDTE;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>Basic implementation supporting XAML files.</summary>
    /// <remarks>This implementation supports resx files using ResXFileCodeGenerator custom tool and has a very low priority so other implementations can be used instead for specific projects.</remarks>
    public class GenericXamlExtractResourceAction : IExtractResourceAction {

        #region IExtractResourceAction Members

        /// <summary>
        /// This action has a very low priority since it is generic to all VB.Net projects.
        /// </summary>
        public virtual int Priority {
            get { return 10; }
        }

        /// <summary>
        /// Gets the default relative path for resource file.
        /// </summary>
        public virtual string DefaultResourceFilePath {
            get { return String.Empty; }
        }

        /// <summary>
        /// Supports all XAML files and all projects
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public virtual bool QuerySupportForProject(EnvDTE.ProjectItem item) {
            return item != null && item.Document.Language.Equals("XAML");
        }

        /// <summary>
        /// This method will be used for filtering resource files displayed to user
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        /// <returns>true if resource file is valid and should be displayed to user</returns>
        public virtual bool IsValidResourceFile(EnvDTE.ProjectItem item) {
            if (item == null) return false;
            try {
                if (item.Properties.Item("Extension").Value.ToString().Equals(".resx", StringComparison.InvariantCultureIgnoreCase)) {

                    if (item.Properties.Item("ItemType").Value.Equals("EmbeddedResource") /*&&
                        (item.Properties.Item("CustomTool").Value.Equals("ResXFileCodeGenerator") ||
                         item.Properties.Item("CustomTool").Value.Equals("VbMyResourcesResXFileCodeGenerator"))*/) {
                        return true;
                    }
                }
                return false;
            } catch (ArgumentException) {
                // This can happen if item does not contain the properties we are looking for.
                return false;
            }
        }

        /// <summary>
        /// Returns the code reference to resource specified in the parameters
        /// </summary>
        /// <param name="file">Resource file containing the resource</param>
        /// <param name="resourceName">Name of the resource</param>
        /// <returns>a piece of code that would reference to the resource provided</returns>
        /// <remarks>This method does not verify if resource actually exists</remarks>
        public string GetResourceReference(ResourceFile file, string resourceName, Project project) {
            string prefix = "";
            if (ExtensibilityMethods.GetProjectType(project) == ProjectType.VB) {
                try {
                    prefix = (string)project.Properties.Item("RootNamespace").Value;
                    if (!string.IsNullOrEmpty(prefix)) prefix += ".";
                } catch { }
            }

            if (file == null) {
                throw new ArgumentNullException("file");
            }
            if (String.IsNullOrEmpty(resourceName)) {
                throw new ArgumentException(Strings.InvalidResourceName, "resourceName");
            }
            string namespacePrefix = GetNamespacePrefix(file);
            if (!file.IsDefaultResXFile()) {
                namespacePrefix += Path.GetFileNameWithoutExtension(file.DisplayName).Replace(' ', '_') + ".";
            } else {
                namespacePrefix += Path.GetFileNameWithoutExtension(file.FileName).Replace(' ', '_') + ".";
            }
            string reference =namespacePrefix + resourceName.Replace(' ', '_');
            return prefix + reference;
        }

        #endregion

        /// <summary>
        /// Determines the namespace of the provided resource file
        /// </summary>
        /// <param name="file">Reference to the resource file</param>
        /// <returns>Namespace to be used to access the resource file</returns>
        protected virtual string GetNamespacePrefix(ResourceFile file) {
            if (file == null) {
                throw new ArgumentNullException("file");
            }
            string namespacePrefix = file.CustomToolNamespace;
            if (String.IsNullOrEmpty(namespacePrefix)) {
                namespacePrefix = file.FileNamespace;
            }
            if (!String.IsNullOrEmpty(namespacePrefix)) {
                namespacePrefix += ":";
            }
            return namespacePrefix;
        }

        #region IExtractResourceAction Members

        /// <summary>
        /// This method should update properties on a recently created resource file so that it
        /// is correctly supported by the same instance of IExtractResourceAction
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        public virtual void UpdateResourceFileProperties(ProjectItem item) {
            if (item != null) {
                item.Properties.Item("CustomTool").Value = "ResXFileCodeGenerator";
                item.Properties.Item("ItemType").Value = "EmbeddedResource";
                //For VB project, we need to set namespace property as well
                item.Properties.Item("CustomToolNamespace").Value = "My.Resources";
            }
        }
        #endregion
    }
}
