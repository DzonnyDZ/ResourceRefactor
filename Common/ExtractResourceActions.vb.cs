﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;
using System.IO;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {

    /// <summary>Basic action implementation supporting all VB projects.</summary>
    /// <remarks>This implementation supports resx files using ResXFileCodeGenerator custom tool and
    /// has a very low priority so other implementations can be used instead for specific projects.</remarks>
    public class GenericVBExtractResourceAction : ExtractResourceActionBase {

        /// <summary>Queries if this action supports the provided project item and its containing project</summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns>True if the action is supported, false otherwise. This implementation supports all VB projects</returns>
        public override bool QuerySupportForProject(EnvDTE.ProjectItem item) {
            if (item == null) return false;
            return
                ExtensibilityMethods.GetProjectType(item.ContainingProject) == ProjectType.VB &&
                item.Document.Language.Equals("Basic");
        }


        /// <summary>Returns the code reference to resource specified in the parameters</summary>
        /// <param name="file">Resource file containing the resource</param>
        /// <param name="resourceName">Name of the resource</param>
        /// <returns>a piece of code that would reference to the resource provided</returns>
        /// <remarks>This method does not verify if resource actually exists</remarks>
        public override string GetResourceReference(ResourceFile file, string resourceName, Project project, BaseHardCodedString @string) {
            if (file == null) {
                throw new ArgumentNullException("file");
            }
            if (String.IsNullOrEmpty(resourceName)) {
                throw new ArgumentException(Strings.InvalidResourceName, "resourceName");
            }
            string namespacePrefix = GetNamespacePrefix(file);
            if (!file.IsDefaultResXFile()) {
                namespacePrefix += Path.GetFileNameWithoutExtension(file.ShortFileName).Replace(' ', '_') + ".";
            }
            string reference = namespacePrefix + resourceName.Replace(' ', '_');
            return reference;
        }

        /// <summary>Determines the namespace of the provided resource file</summary>
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
                namespacePrefix += ".";
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
                //For VB project, we need to set namespace property as well
                item.Properties.Item("CustomToolNamespace").Value = "My.Resources";
            }
        }
    }

    /// <summary>Resource extration acction implementation for Visual Basic Razor files (VBHTML)</summary>
    public class VBRazorExtractResourceAction : GenericVBExtractResourceAction {
        /// <summary>Queries if this action supports the provided project item and its containing project</summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns>True if the action is supported, false otherwise. This implementation supports Rator (VBHTML) fiels in VB projects</returns>
        public override bool QuerySupportForProject(ProjectItem item) {
            if (item == null) return false;
            return
                ExtensibilityMethods.GetProjectType(item.ContainingProject) == ProjectType.VB
                && item.Document.Language.Equals("HTML") && item.Document.Name.EndsWith(".vbhtml", StringComparison.CurrentCultureIgnoreCase); ;
        }
    }

    /// <summary>Implementation supporting VB file and website projects.</summary>
    /// <remarks>This implementation is used when string is in a VB file which is contained by a website project</remarks>
    public class WebsiteVBExtractResourceAction : GenericVBExtractResourceAction {

        /// <summary>Gets priority of the action.</summary>
        public override int Priority { get { return 50; } }

        /// <summary>Gets the default relative path for resource file.</summary>
        public override string DefaultResourceFilePath { get { return "App_GlobalResources"; } }

        /// <summary>Queries if this action supports the provided project item and its containing project</summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns></returns>
        public override bool QuerySupportForProject(ProjectItem item) {
            return
                item != null && item.Document.Language.Equals("Basic") &&
                ExtensibilityMethods.GetProjectType(item.ContainingProject) == ProjectType.WebProject;
        }

        /// <summary>
        /// This method will be used for filtering resource files displayed to user
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        /// <returns>true if resource file is valid and should be displayed to user</returns>
        public override bool IsValidResourceFile(ProjectItem item) {
            return WebsiteCSharpExtractResourceAction.CheckResourceFileForWebSites(item);
        }

        /// <summary>Determines the namespace of the provided resource file</summary>
        /// <param name="file">Reference to the resource file</param>
        /// <returns>Namespace to be used to access the resource file</returns>
        protected override string GetNamespacePrefix(ResourceFile file) {
            return "Resources.";
        }

        /// <summary>
        /// This method should update properties on a recently created resource file so that it
        /// is correctly supported by the same instance of <see cref="IExtractResourceAction"/>
        /// </summary>
        /// <param name="item">Project item for the resource file</param>
        /// <remarks>This implementation does nothing</remarks>
        public override void UpdateResourceFileProperties(ProjectItem item) {}
    }

    /// <summary>Resource extract implementation supporting VB.Net file and web application projects.</summary>
    public class WebApplicationVBExtractResourceAction : GenericVBExtractResourceAction {
        /// <summary>Gets priority of the action.</summary>
        public override int Priority { get { return 50; } }

        /// <summary>Gets the default relative path for resource file.</summary>
        public override string DefaultResourceFilePath { get { return "App_GlobalResources"; } }

        /// <summary>Queries if this action supports the provided project item and its containing project</summary>
        /// <param name="item">Project item to query support for</param>
        /// <returns>True if the action is supported, false otherwise. This implementation supports web applications</returns>
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

        /// <summary>This method will be used for filtering resource files displayed to user</summary>
        /// <param name="item">Project item for the resource file</param>
        /// <returns>true if resource file is valid and should be displayed to user</returns>
        public override bool IsValidResourceFile(ProjectItem item) {
            return base.IsValidResourceFile(item) || WebApplicationCSharpExtractResourceAction.IsValidContentResourceFile(item);
        }

        /// <summary>Determines the namespace of the provided resource file</summary>
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
                string relPath = WebApplicationCSharpExtractResourceAction.GetRelativePathForItem(item);
                if (relPath.Equals("App_GlobalResources")) {
                    item.Properties.Item("CustomTool").Value = "GlobalResourceProxyGenerator";
                    item.Properties.Item("ItemType").Value = "Content";
                } else {
                    base.UpdateResourceFileProperties(item);
                }
            }
        }
    }
}
