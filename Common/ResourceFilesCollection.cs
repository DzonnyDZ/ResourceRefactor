// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Windows.Forms;
using System.IO;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common
{

    /// <summary>Represents a collection of resource files in a Visual Studio project.</summary>
    /// <remarks>Collection is represented as both a list form and a TreeNode structure.</remarks>
    public class ResourceFileCollection : FilteredProjectCollection<ResourceFile>
    {

        /// <summary>Additional filtering method to filter resource file entries.</summary>
        private FilterMethod supplementFilter;

        #region Public Properties

        /// <summary>Gets the resource file with the provided display name.</summary>
        /// <param name="displayName">Display name to get resource file for</param>
        /// <returns>Resource file for <paramref name="displayName"/></returns>
        /// <exception cref="ArgumentOutOfRangeException">Resource file with given <paramref name="displayName"/> is not present in this collection</exception>
        public ResourceFile this[string displayName]
        {
            get
            {
                foreach (ResourceFile file in this)
                {
                    if (file.DisplayName.Equals(displayName))
                        return file;
                }
                throw new ArgumentOutOfRangeException();
            }
        }

        #endregion

        /// <summary>Creates a new collection from provided project, collection will contain all resource files in the project</summary>
        /// <param name="project">Project to load resource files from</param>
        /// <param name="filter">Filtering method to decide whether to include a file or not</param>
        public ResourceFileCollection(Project project, FilterMethod filter) : base(project)
        {
            this.FilteringMethod = new FilterMethod(this.IsValidResource);
            this.supplementFilter = filter;
            this.RefreshListOfFiles();
        }

        /// <summary>Gets the first instance of resource file with the provided display name.</summary>
        /// <param name="displayName">Display name to get resource file with</param>
        /// <returns>ResourceFile if found, null otherwise</returns>
        public ResourceFile GetResourceFile(string displayName)
        {
            foreach (ResourceFile file in this)
            {
                if (file.DisplayName.Equals(displayName))
                    return file;
            }
            return null;
        }

        /// <summary>Adds a project item to the list by wrapping it as a Resource File.</summary>
        /// <param name="item">An item to be added</param>
        public override void AddItem(ProjectItem item)
        {
            this.Add(new ResourceFile(item));
        }

        /// <summary>Checks if a project item is valid and file associated with it exists in the system, also calls the supplemental filter method if one is avaiable</summary>
        /// <param name="item">An item fo test</param>
        /// <returns>True if <paramref name="item"/> is valid, the file exists and filter allows it; false otherwise</returns>
        private bool IsValidResource(ProjectItem item)
        {
            bool result = false;
            if (item != null && File.Exists(item.get_FileNames(0))) 
            {
                if (this.supplementFilter != null)
                {
                    result = this.supplementFilter(item);
                }
                else
                {
                    result = true;
                }
            }
            return result;
        }
    }
}
