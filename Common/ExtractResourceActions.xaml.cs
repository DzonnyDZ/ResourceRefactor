using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using EnvDTE;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>Basic implementation or resource extract action supporting XAML files.</summary>
    public class GenericXamlExtractResourceAction : ExtractResourceActionBase {

        /// <summary>Queries if this action supports the provided project item and its containing project</summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns>True if the action is supported, false otherwise. This implementation supports XAML files.</returns>
        public override bool QuerySupportForProject(EnvDTE.ProjectItem item) {
            return item != null && item.Document.Language.Equals("XAML");
        }

        /// <summary>Returns the code reference to resource specified in the parameters</summary>
        /// <param name="file">Resource file containing the resource</param>
        /// <param name="resourceName">Name of the resource</param>
        /// <param name="project">Project current file belongs to</param>
        /// <param name="string">String being extracted</param>
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
                namespacePrefix += Path.GetFileNameWithoutExtension(file.ShortFileName).Replace(' ', '_') + ".";
            } else {
                namespacePrefix += Path.GetFileNameWithoutExtension(file.FileName).Replace(' ', '_') + ".";
            }
            string reference = namespacePrefix + resourceName.Replace(' ', '_');
            return prefix + reference;
        }


        /// <summary>Retruns namespace prefix for a given file</summary>
        /// <param name="file">A file to get namespace prefix for</param>
        /// <returns>Namespace prefix for <paramref name="file"/></returns>
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

        /// <summary>This method should update properties on a recently created resource file so that it is correctly supported by the same instance of <see cref="IExtractResourceAction"/></summary>
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
        /// <summary>Gets priority of the action.</summary>
        public override int Priority {
            get {
                return base.Priority + 1;
            }
        }

        /// <summary>Queries if this action supports the provided project item and its containing project</summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns>True if the action is supported, false otherwise. This implementation supports XAML files in VB projects.</returns>
        public override bool QuerySupportForProject(EnvDTE.ProjectItem item) {
            if (item == null) return false;
            return item != null && item.Document.Language.Equals("XAML") &&
                ExtensibilityMethods.GetProjectType(item.ContainingProject) == ProjectType.VB &&
                item.Document.Language.Equals("Basic");
        }

        /// <summary>This method should update properties on a recently created resource file so that it is correctly supported by the same instance of <see cref="IExtractResourceAction"/></summary>
        /// <param name="item">Project item for the resource file</param>
        public override void UpdateResourceFileProperties(ProjectItem item) {
            base.UpdateResourceFileProperties(item);
            if (item != null)
                //For VB project, we need to set namespace property as well
                item.Properties.Item("CustomToolNamespace").Value = "My.Resources";
        }
    }
}
