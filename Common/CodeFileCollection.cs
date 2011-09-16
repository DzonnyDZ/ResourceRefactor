/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common
{
    /// <summary>
    /// Collection of code files in a Visual Studio project. Depending on type of the project, this will only include
    /// *.cs or *.vb files (or both). Also all designer generated files are excluded from the collection.
    /// </summary>
    public class CodeFileCollection : FilteredProjectCollection<ProjectItem>
    {
        /// <summary>
        /// Determines what type of code files to list in the collection
        /// </summary>
        [Flags]
        enum CodeType
        {
            None = 0,
            CSharp = 1,
            VB = 2,
            Both = CSharp | VB
        }

        /// <summary>
        /// Code type filtering used when recursing in to project tree.
        /// </summary>
        private CodeType codeTypeFilter;

        /// <summary>
        /// Gets the first instance of resource file with the provided display name.
        /// </summary>
        /// <param name="displayName"></param>
        /// <returns>ResourceFile if found, null otherwise</returns>
        public ProjectItem GetCodeFile(string displayName)
        {
            foreach (ProjectItem item in this)
            {
                if (item.Name.Equals(displayName))
                {
                    return item;
                }
            }
            return null;
        }

        /// <summary>
        /// Creates a new code file collection that lists all the code files in a project that can be safely edited (Designer 
        /// code files are not excluded)
        /// </summary>
        /// <param name="project">Project to list code files</param>
        public CodeFileCollection(Project project) : base(project, null)
        {
            ProjectType type = ExtensibilityMethods.GetProjectType(project);
            switch (type)
            {
                case ProjectType.CSharp:
                    this.codeTypeFilter = CodeType.CSharp;
                    break;
                case ProjectType.VB:
                    this.codeTypeFilter = CodeType.VB;
                    break;
                case ProjectType.WebProject:
                    this.codeTypeFilter = CodeType.Both;
                    break;
                default:
                    throw new ArgumentException(Strings.ProjectFileInvalid, "project");
            }
            this.FilteringMethod = new FilterMethod(this.IsValidCodeFile);
            this.RefreshListOfFiles();
       }

        /// <summary>
        /// Checks if an item is a valid code file for the type of parent project (Designer files are not included 
        /// in this collection since they should not be modified)
        /// </summary>
        /// <param name="item">ProjectI</param>
        /// <returns></returns>
        private bool IsValidCodeFile(ProjectItem item)
        {
            bool result = false;
            try 
            {
                if ((this.codeTypeFilter & CodeType.CSharp) == CodeType.CSharp)
                {
                    if (!item.Name.EndsWith(".Designer" + Strings.ExtensionCSharp))
                    {
                        result = result || item.Properties.Item("Extension").Value.ToString().Equals(Strings.ExtensionCSharp);
                    }
                }
                if (!result && ((this.codeTypeFilter & CodeType.VB) == CodeType.VB))
                {
                    if (!item.Name.EndsWith(".Designer" + Strings.ExtensionVB))
                    {
                        result = result || item.Properties.Item("Extension").Value.ToString().Equals(Strings.ExtensionVB);
                    }
                }
                return result;
            } 
            catch (ArgumentException) 
            {
                return false;
            }
        }
    }
}
