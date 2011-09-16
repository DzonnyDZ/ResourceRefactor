/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Windows.Forms;
using System.IO;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common
{

    /// <summary>
    /// Represents a collection of resource files in a Visual Studio project. Collection is represented as
    /// both a list form and a TreeNode structure.
    /// </summary>
    public class ResourceFileCollection : FilteredProjectCollection<ResourceFile>
    {

        /// <summary>
        /// Additional filtering method to filter resource file entries.
        /// </summary>
        private FilterMethod supplementFilter;

        #region Public Properties

        /// <summary>
        /// Gets the resource file with the provided display name.
        /// </summary>
        /// <param name="displayName"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Creates a new collection from provided project, collection will contain
        /// all resource files in the project
        /// </summary>
        /// <param name="project"></param>
        public ResourceFileCollection(Project project, FilterMethod filter) : base(project)
        {
            this.FilteringMethod = new FilterMethod(this.IsValidResource);
            this.supplementFilter = filter;
            this.RefreshListOfFiles();
        }

        /// <summary>
        /// Gets the first instance of resource file with the provided display name.
        /// </summary>
        /// <param name="displayName"></param>
        /// <returns>ResourceFile if found, null otherwise</returns>
        public ResourceFile GetResourceFile(string displayName)
        {
            foreach (ResourceFile file in this)
            {
                if (file.DisplayName.Equals(displayName))
                {
                    return file;
                }
            }
            return null;
        }

        /// <summary>
        /// Adds a project item to the list by wrapping it as a Resource File.
        /// </summary>
        /// <param name="item"></param>
        public override void AddItem(ProjectItem item)
        {
            this.Add(new ResourceFile(item));
        }

        /// <summary>
        /// Checks if a project item is valid and file associated with it exists in the system,
        /// also calls the supplemental filter method if one is avaiable
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
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
