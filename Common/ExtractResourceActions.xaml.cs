using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using EnvDTE;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>Basic implementation supporting XAML files.</summary>
    /// <remarks>This implementation supports resx files using ResXFileCodeGenerator custom tool and has a very low priority so other implementations can be used instead for specific projects.</remarks>
    public class GenericXamlExtractResourceAction : ExtractResourceActionBase {

        /// <summary>Supports all XAML files and all projects</summary>
        public override bool QuerySupportForProject(EnvDTE.ProjectItem item) {
            return item != null && item.Document.Language.Equals("XAML");
        }

        /// <summary>
        /// Returns the code reference to resource specified in the parameters
        /// </summary>
        /// <param name="file">Resource file containing the resource</param>
        /// <param name="resourceName">Name of the resource</param>
        /// <returns>a piece of code that would reference to the resource provided</returns>
        /// <remarks>This method does not verify if resource actually exists</remarks>
        public override string GetResourceReference(ResourceFile file, string resourceName, Project project, BaseHardCodedString @string) {
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
            string reference = namespacePrefix + resourceName.Replace(' ', '_');
            return prefix + reference;
        }


        /// <summary>
        /// Determines the namespace of the provided resource file
        /// </summary>
        /// <param name="file">Reference to the resource file</param>
        /// <returns>Namespace to be used to access the resource file</returns>
        protected override string GetNamespacePrefix(ResourceFile file) {
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

        /// <summary>
        /// This method should update properties on a recently created resource file so that it
        /// is correctly supported by the same instance of IExtractResourceAction
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        public override void UpdateResourceFileProperties(ProjectItem item) {
            if (item != null) {
                item.Properties.Item("CustomTool").Value = "ResXFileCodeGenerator";
                item.Properties.Item("ItemType").Value = "EmbeddedResource";
            }
        }
    }

    /// <summary>Special implementation of XAML resource extractor for Visual Basic</summary>
    public class VbExtractResourceXamlAction : GenericXamlExtractResourceAction {
        public override int Priority {
            get {
                return base.Priority + 1;
            }
        }
        public override bool QuerySupportForProject(EnvDTE.ProjectItem item) {
            if (item == null) return false;
            return item != null && item.Document.Language.Equals("XAML") &&
                ExtensibilityMethods.GetProjectType(item.ContainingProject) == ProjectType.VB &&
                item.Document.Language.Equals("Basic");
        }
        /// <summary>
        /// This method should update properties on a recently created resource file so that it
        /// is correctly supported by the same instance of IExtractResourceAction
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        public override void UpdateResourceFileProperties(ProjectItem item) {
            base.UpdateResourceFileProperties(item);
            if (item != null)
                //For VB project, we need to set namespace property as well
                item.Properties.Item("CustomToolNamespace").Value = "My.Resources";
        }
    }
}
