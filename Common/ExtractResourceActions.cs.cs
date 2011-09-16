using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;
using System.IO;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>
    /// Basic implementation supporting all C# projects.
    /// </summary>
    /// <remarks>This implementation supports resx files using ResXFileCodeGenerator custom tool and
    /// has a very low priority so other implementations can be used instead for specific projects.</remarks>
    public class GenericCSharpExtractResourceAction : IExtractResourceAction {

        #region IExtractResourceAction Members

        /// <summary>
        /// This action has a very low priority since it is generic to all C# projects.
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
        /// Supports all C# files and C# projects
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public virtual bool QuerySupportForProject(EnvDTE.ProjectItem item) {
            if (item == null) return false;
            return
                item.Document.Language.Equals("CSharp") &&
                ExtensibilityMethods.GetProjectType(item.ContainingProject) == ProjectType.CSharp;
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
                    (item.Properties.Item("CustomTool").Value.Equals("ResXFileCodeGenerator"))*/) {
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
        /// Determines the namespace of the provided resource file
        /// </summary>
        /// <param name="file">Reference to the resource file</param>
        /// <returns>Namespace to be used to access the resource file</returns>
        protected virtual string GetNamespacePrefix(ResourceFile file) {
            string namespacePrefix = String.Empty;
            try {
                namespacePrefix = file.CustomToolNamespace;
                if (String.IsNullOrEmpty(namespacePrefix) && file.FileNamespace != null) {
                    namespacePrefix = file.FileNamespace;
                } else {
                    namespacePrefix = file.Item.ContainingProject.Properties.Item("DefaultNamespace").Value.ToString();
                }
                namespacePrefix += ".";
            } catch (ArgumentException) {
                namespacePrefix = String.Empty;
            }
            return namespacePrefix;
        }

        /// <summary>
        /// Returns the code reference to resource specified in the parameters
        /// </summary>
        /// <param name="file">Resource file containing the resource</param>
        /// <param name="resourceName">Name of the resource</param>
        /// <returns>a piece of code that would reference to the resource provided</returns>
        /// <remarks>This method does not verify if resource actually exists</remarks>
        public virtual string GetResourceReference(ResourceFile file, string resourceName, Project project) {
            if (file == null) {
                throw new ArgumentNullException("file");
            }
            if (String.IsNullOrEmpty(resourceName)) {
                throw new ArgumentException(Strings.InvalidResourceName, "resourceName");
            }
            string namespacePrefix = this.GetNamespacePrefix(file);
            string reference = namespacePrefix + Path.GetFileNameWithoutExtension(file.DisplayName).Replace(' ', '_') + "." + resourceName.Replace(' ', '_');
            return reference;
        }

        /// <summary>
        /// This method should update properties on a recently created resource file so that it
        /// is correctly supported by the same instance of IExtractResourceAction
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        public virtual void UpdateResourceFileProperties(ProjectItem item) {
            if (item != null) {
                item.Properties.Item("CustomTool").Value = "ResXFileCodeGenerator";
                item.Properties.Item("ItemType").Value = "EmbeddedResource";
            }
        }

        #endregion
    }


    /// <summary>
    /// Implementation supporting C# file and website projects.
    /// </summary>
    /// <remarks>This implementation is used when string is in a C# file which is contained by a website project</remarks>
    public class WebsiteCSharpExtractResourceAction : GenericCSharpExtractResourceAction {
        #region IExtractResourceAction Members

        /// <summary>
        /// Priority of the action. If there are multiple actions supporting the same item, action with
        /// the highest priority will be selected
        /// </summary>
        public override int Priority {
            get { return 50; }
        }

        /// <summary>
        /// Gets the default relative path for resource file.
        /// </summary>
        public override string DefaultResourceFilePath {
            get {
                return "App_GlobalResources";
            }
        }

        /// <summary>
        /// Queries if this action supports the provided project item and its containing project
        /// </summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns></returns>
        public override bool QuerySupportForProject(ProjectItem item) {
            return
                item != null && item.Document.Language.Equals("CSharp") &&
                ExtensibilityMethods.GetProjectType(item.ContainingProject) == ProjectType.WebProject;
        }

        /// <summary>
        /// This method will be used for filtering resource files displayed to user
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        /// <returns>true if resource file is valid and should be displayed to user</returns>
        public override bool IsValidResourceFile(ProjectItem item) {
            return CheckResourceFileForWebSites(item);
        }

        #endregion

        /// <summary>
        /// Determines the namespace of the provided resource file
        /// </summary>
        /// <param name="file">Reference to the resource file</param>
        /// <returns>Namespace to be used to access the resource file</returns>
        protected override string GetNamespacePrefix(ResourceFile file) {
            return "Resources.";
        }

        /// <summary>
        /// Checks if provided item is a valid resource file for website projects
        /// </summary>
        /// <param name="item"></param>
        /// <returns>true if item is a valid resource file</returns>
        internal static bool CheckResourceFileForWebSites(ProjectItem item) {
            try {
                if (item.Properties.Item("Extension").Value.ToString().Equals(".resx", StringComparison.InvariantCultureIgnoreCase)) {
                    return Path.GetDirectoryName(item.Properties.Item("RelativeURL").Value.ToString()).Equals("App_GlobalResources");
                }
                return false;
            } catch (ArgumentException) {
                // This can happen if item does not contain the properties we are looking for.
                return false;
            }
        }

        /// <summary>
        /// This method should update properties on a recently created resource file so that it
        /// is correctly supported by the same instance of IExtractResourceAction
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        public override void UpdateResourceFileProperties(ProjectItem item) {
        }
    }


    /// <summary>
    /// Implementation supporting C# file and web application projects.
    /// </summary>
    public class WebApplicationCSharpExtractResourceAction : GenericCSharpExtractResourceAction {
        #region IExtractResourceAction Members

        /// <summary>
        /// Priority of the action. If there are multiple actions supporting the same item, action with
        /// the highest priority will be selected
        /// </summary>
        public override int Priority {
            get { return 50; }
        }

        /// <summary>
        /// Gets the default relative path for resource file.
        /// </summary>
        public override string DefaultResourceFilePath {
            get {
                return "App_GlobalResources";
            }
        }

        /// <summary>
        /// Queries if this action supports the provided project item and its containing project
        /// </summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns></returns>
        public override bool QuerySupportForProject(ProjectItem item) {
            if (item == null) return false;
            bool returnValue = base.QuerySupportForProject(item);
            bool extenderFound = false;
            if (returnValue) {
                string[] names = (string[])(item.ContainingProject.ExtenderNames);
                foreach (string name in names) {
                    if (name.Equals("WebApplication")) {
                        extenderFound = true;
                        break;
                    }
                }
            }
            return returnValue && extenderFound;
        }

        /// <summary>
        /// This method will be used for filtering resource files displayed to user
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        /// <returns>true if resource file is valid and should be displayed to user</returns>
        public override bool IsValidResourceFile(ProjectItem item) {
            return base.IsValidResourceFile(item) || IsValidContentResourceFile(item);
        }

        #endregion

        /// <summary>
        /// Determines the namespace of the provided resource file
        /// </summary>
        /// <param name="file">Reference to the resource file</param>
        /// <returns>Namespace to be used to access the resource file</returns>
        protected override string GetNamespacePrefix(ResourceFile file) {
            if (file.CustomToolName.Equals("GlobalResourceProxyGenerator")) {
                return "Resources.";
            } else {
                return base.GetNamespacePrefix(file);
            }
        }

        /// <summary>
        /// This method should update properties on a recently created resource file so that it
        /// is correctly supported by the same instance of IExtractResourceAction
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        public override void UpdateResourceFileProperties(ProjectItem item) {
            if (item != null) {
                string relPath = GetRelativePathForItem(item);
                if (relPath.Equals("App_GlobalResources")) {
                    item.Properties.Item("CustomTool").Value = "GlobalResourceProxyGenerator";
                    item.Properties.Item("ItemType").Value = "Content";
                } else {
                    base.UpdateResourceFileProperties(item);
                }
            }
        }

        /// <summary>
        /// Checks if resource file is a valid content resource file used in web application projects
        /// These resource files must be located in App_GlobalResources directory
        /// </summary>
        /// <returns></returns>
        public static bool IsValidContentResourceFile(ProjectItem item) {
            if (item == null) return false;
            try {
                if (item.Properties.Item("Extension").Value.ToString().Equals(".resx", StringComparison.InvariantCultureIgnoreCase)) {
                    if (item.Properties.Item("ItemType").Value.Equals("Content") &&
                        item.Properties.Item("CustomTool").Value.Equals("GlobalResourceProxyGenerator")) {
                        string relPath = GetRelativePathForItem(item);
                        return relPath.Equals("App_GlobalResources");
                    }
                }
                return false;
            } catch (ArgumentException) {
                // This can happen if item does not contain the properties we are looking for.
                return false;
            }
        }

        /// <summary>
        /// Gets the relative path of the item in the project
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static string GetRelativePathForItem(ProjectItem item) {
            if (item == null) {
                throw new ArgumentNullException("item");
            }
            string itemFullPath = Path.GetDirectoryName(item.get_FileNames(1));
            string projectFullPath = Path.GetDirectoryName(item.ContainingProject.FileName);
            if (itemFullPath.StartsWith(projectFullPath) && itemFullPath.Length > projectFullPath.Length) {
                int increment = (projectFullPath.EndsWith(@"\")) ? 0 : 1;
                return itemFullPath.Substring(projectFullPath.Length + increment);
            } else {
                return String.Empty;
            }
        }
    }


}
