/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Windows.Forms;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {

    /// <summary>Delegate used by <see cref="FilteredProjectCollection"/> to filter elements</summary>
    /// <param name="item">An item to test</param>
    /// <returns>True to include the item, false not to include</returns>
    public delegate bool FilterMethod(ProjectItem item);

    /// <summary>Represents a collection of resource files in a Visual Studio project.</summary>
    /// <typeparam name="T">Type of items in colletion</typeparam>
    /// <remarks>Collection is represented as both a list form and a TreeNode structure.</remarks>
    public class FilteredProjectCollection<T> : System.Collections.ObjectModel.Collection<T> {
        private readonly Project project;
        private TreeNode rootNode;

        #region Public Properties

        /// <summary>Gets root tree node to be used by list boxes</summary>
        public TreeNode RootNode { get { return rootNode; } }

        /// <summary>Gets root project object used to create this collection.</summary>
        public Project Project { get { return project; } }

        /// <summary>Gets or sets filtering method to be used when recursing through the project.</summary>
        /// <remarks>If the method is specified and it returns false, item will not be included in the collection</remarks>
        public FilterMethod FilteringMethod { get; set; }
        #endregion

        /// <summary>Creates a new collection from provided project, collection will contain all resource files in the project</summary>
        /// <param name="project">Root project object used to create this collection</param>
        /// <param name="filteringMethod">Filtering method to be used when recursing through the project. If the method is specified and it returns false, item will not be included in the collection</param>
        public FilteredProjectCollection(Project project, FilterMethod filteringMethod) {
            this.project = project;
            this.FilteringMethod = filteringMethod;
            this.rootNode = CreateProjectFileTree(this.project);
        }

        /// <summary>
        /// Creates a new collection from provided project, collection will contain
        /// all resource files in the project
        /// </summary>
        /// <param name="project"></param>
        /// <remarks>RefreshListOfFiles should be called by the caller since filtering method is provided later</remarks>
        public FilteredProjectCollection(Project project) {
            this.project = project;
        }

        /// <summary>Refreshes resource file list from the root project object</summary>
        public void RefreshListOfFiles() {
            this.Clear();
            this.rootNode = CreateProjectFileTree(this.project);
        }

        /// <summary>Adds item to the list.</summary>
        /// <param name="item">Project item to add</param>
        /// <remarks>
        /// Default behaviour is to try cast the item to type T and add it
        /// <note type="inheritinfo">Classes inheriting this class should override this method for correct behaviour.</note>
        /// </remarks>
        public virtual void AddItem(ProjectItem item) {
            this.Add((T)item);
        }

        #region Private Methods

        /// <summary>Creates a tree of files included in the project.</summary>
        /// <param name="project">Visual Studio project object</param>
        /// <returns>A collection of TreeNode object where each node is a sub project or a file entry</returns>
        private TreeNode CreateProjectFileTree(Project project) {
            TreeNode parent = new TreeNode();
            parent.Text = project.Name;
            parent.Name = parent.Text;
            parent.Tag = project;
            FillProjectFileTree(project.ProjectItems, parent);
            return parent;
        }

        /// <summary>Creates a tree of files included in the project.</summary>
        /// <param name="project">Visual Studio project object</param>
        /// <returns>A collection of TreeNode object where each node is a sub project or a file entry</returns>
        private TreeNode CreateProjectFileTree(ProjectItem projectItem) {
            TreeNode parent = new TreeNode();
            parent.Name = projectItem.Name;
            parent.Text = parent.Name;
            parent.Tag = projectItem;
            FillProjectFileTree(projectItem.ProjectItems, parent);
            return parent;
        }

        /// <summary>Creates child nodes under a parent node form a ProjectItems collection</summary>
        /// <param name="items">Project item collection</param>
        /// <param name="parent">Parent node to fill</param>
        private void FillProjectFileTree(ProjectItems items, TreeNode parent) {
            foreach (ProjectItem item in items) {
                TreeNode node = new TreeNode();
                node.Name = item.Name;
                node.Text = item.Name;
                node.Tag = item;
                if (item.ProjectItems != null && item.ProjectItems.Count > 0) {
                    node = CreateProjectFileTree(item);
                    /// If node has no children and fails the filter do not add it.
                    if (node.Nodes.Count == 0 && (FilteringMethod != null && !FilteringMethod(item))) {
                        continue;
                    }
                }
                parent.Nodes.Add(node);
                if (FilteringMethod == null || FilteringMethod(item)) {
                    this.AddItem(item);
                }
            }
        }

        #endregion
    }
}
