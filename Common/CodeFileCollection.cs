/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common
{
    /// <summary>Collection of code files in a Visual Studio project.</summary>
    /// <remarks>Depending on type of the project, this will only include *.cs or *.vb files (or both). Also all designer generated files are excluded from the collection.</remarks>
    public class CodeFileCollection : FilteredProjectCollection<ProjectItem>
    {
        /// <summary>Determines what type of code files to list in the collection</summary>
        [Flags]
        enum CodeType
        {
            /// <summary>Include no files</summary>
            None = 0,
            /// <summary>Include C# files</summary>
            CSharp = 1,
            /// <summary>Include Visual Basic files</summary>
            VB = 2,
            /// <summary>Inlcude C# and Visual Basic files</summary>
            Both = CSharp | VB
        }

        /// <summary>Code type filtering used when recursing in to project tree.</summary>
        private CodeType codeTypeFilter;

        /// <summary>Gets the first instance of resource file with the provided display name.</summary>
        /// <param name="displayName">A display nime to get resource ofr</param>
        /// <returns>ResourceFile if found, null otherwise</returns>
        public ProjectItem GetCodeFile(string displayName)
        {
            foreach (ProjectItem item in this)
            {
                if (item.Name.Equals(displayName))
                    return item;
            }
            return null;
        }

        /// <summary>Creates a new code file collection that lists all the code files in a project that can be safely edited</summary>
        /// <param name="project">Project to list code files</param>
        /// <remarks>Designer code files are not excluded</remarks>
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

        /// <summary>Checks if an item is a valid code file for the type of parent project</summary>
        /// <param name="item">ProjectI</param>
        /// <returns>True if <paramref name="item"/> is a valid code file for it's parent project</returns>
        /// <remarks>Designer files are not included in this collection since they should not be modified</remarks>
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