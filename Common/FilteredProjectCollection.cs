/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Windows.Forms;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common
{

    /// <summary>
    /// Delegate used by FilteredProjectCollection to filter elements
    /// </summary>
    /// <param name="item"></param>
    /// <returns></returns>
    public delegate bool FilterMethod(ProjectItem item);

    /// <summary>
    /// Represents a collection of resource files in a Visual Studio project. Collection is represented as
    /// both a list form and a TreeNode structure.
    /// </summary>
    public class FilteredProjectCollection<T> : System.Collections.ObjectModel.Collection<T>
    {
        
        #region Private Members
        
        /// <summary>
        /// Root project used to create this collection
        /// </summary>
        private Project rootProject;
        
        /// <summary>
        /// Root node for Tree structure
        /// </summary>
        private TreeNode rootNode;

        /// <summary>
        /// Filtering method to use (if null all items are included in the list)
        /// </summary>
        private FilterMethod filterMethod;
        #endregion

        #region Public Properties

        /// <summary>
        /// Root tree node to be used by list boxes
        /// </summary>
        public TreeNode RootNode
        {
            get
            {
                return rootNode;
            }
        }

        /// <summary>
        /// Root project object used to create this collection.
        /// </summary>
        public Project Project
        {
            get
            {
                return this.rootProject;
            }
        }

        /// <summary>
        /// Filtering method to be used when recursing through the project. If the method is specified and it returns false, item
        /// will not be included in the collection
        /// </summary>
        public FilterMethod FilteringMethod
        {
            get
            {
                return this.filterMethod;
            }
            set
            {
                this.filterMethod = value;
            }
        }
        #endregion

        /// <summary>
        /// Creates a new collection from provided project, collection will contain
        /// all resource files in the project
        /// </summary>
        /// <param name="project"></param>
        public FilteredProjectCollection(Project project, FilterMethod filterMethod)
        {
            this.rootProject = project;
            this.filterMethod = filterMethod;
            this.rootNode = CreateProjectFileTree(this.rootProject);
        }

        /// <summary>
        /// Creates a new collection from provided project, collection will contain
        /// all resource files in the project
        /// </summary>
        /// <param name="project"></param>
        /// <remarks>RefreshListOfFiles should be called by the caller since filtering method is provided later</remarks>
        public FilteredProjectCollection(Project project)
        {
            this.rootProject = project;
        }

        /// <summary>
        /// Refreshes resource file list from the root project object
        /// </summary>
        public void RefreshListOfFiles()
        {
            this.Clear();
            this.rootNode = CreateProjectFileTree(this.rootProject);
        }

        /// <summary>
        /// Adds item to the list. Classes inheriting this class should override this method for correct behaviour.
        /// Default behaviour is to try cast the item to type T and add it
        /// </summary>
        /// <param name="item">Project item to add</param>
        public virtual void AddItem(ProjectItem item)
        {
            this.Add((T)item);
        }

        #region Private Methods

        /// <summary>
        /// Creates a tree of files included in the project.
        /// </summary>
        /// <param name="project">Visual Studio project object</param>
        /// <returns>A collection of TreeNode object where each node is a sub project or a file entry</returns>
        private TreeNode CreateProjectFileTree(Project project)
        {
            TreeNode parent = new TreeNode();
            parent.Text = project.Name;
            parent.Name = parent.Text;
            parent.Tag = project;
            FillProjectFileTree(project.ProjectItems, parent);
            return parent;
        }

        /// <summary>
        /// Creates a tree of files included in the project.
        /// </summary>
        /// <param name="project">Visual Studio project object</param>
        /// <returns>A collection of TreeNode object where each node is a sub project or a file entry</returns>
        private TreeNode CreateProjectFileTree(ProjectItem projectItem)
        {
            TreeNode parent = new TreeNode();
            parent.Name = projectItem.Name;
            parent.Text = parent.Name;
            parent.Tag = projectItem;
            FillProjectFileTree(projectItem.ProjectItems, parent);
            return parent;
        }

        /// <summary>
        /// Creates child nodes under a parent node form a ProjectItems collection
        /// </summary>
        /// <param name="items">Project item collection</param>
        /// <param name="parent">Parent node to fill</param>
        private void FillProjectFileTree(ProjectItems items, TreeNode parent)
        {
            foreach (ProjectItem item in items)
            {
                TreeNode node = new TreeNode();
                node.Name = item.Name;
                node.Text = item.Name;
                node.Tag = item;
                if (item.ProjectItems != null && item.ProjectItems.Count > 0)
                {
                    node = CreateProjectFileTree(item);
                    /// If node has no children and fails the filter do not add it.
                    if (node.Nodes.Count == 0 && (this.filterMethod != null && !this.filterMethod(item)))
                    {
                        continue;
                    }
                }
                parent.Nodes.Add(node);
                if (this.filterMethod == null || this.filterMethod(item))
                {
                    this.AddItem(item);
                }
            }
        }

        #endregion
    }
}
