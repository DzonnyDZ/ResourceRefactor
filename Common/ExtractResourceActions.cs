using System;
using System.Data;
using System.Configuration;
using System.IO;
using EnvDTE;
using System.Collections.Generic;
using System.Reflection;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {



    /// <summary>
    /// Represents a region of the code that is going to be refactored.
    /// </summary>
    /// <remarks>Contains methods to preview changes to be applied and to do the actual refactoring</remarks>
    public class ExtractToResourceActionSite {
        private BaseHardCodedString stringInstance;

        private IExtractResourceAction actionObject;

        /// <summary>
        /// Gets the string instance to refactor
        /// </summary>
        public BaseHardCodedString StringToExtract {
            get { return stringInstance; }
        }

        /// <summary>
        /// Gets the action object to use for refactoring
        /// </summary>
        public IExtractResourceAction ActionObject {
            get { return this.actionObject; }
        }

        /// <summary>
        /// Creates a new instance to refactor provided string instance
        /// </summary>
        /// <param name="stringInstance">String instance to refactor</param>
        public ExtractToResourceActionSite(BaseHardCodedString stringInstance) {
            this.stringInstance = stringInstance;
            actionObject = GetExtractResourceActionObject(this.StringToExtract);
        }

        /// <summary>
        /// Returns a preview of the line after the literal is replaced by a reference to the resource
        /// </summary>
        /// <param name="file">Resource file containing the resource</param>
        /// <param name="resourceName">Name of the resource</param>
        /// <returns>Full line where the text would be replaced</returns>
        public string PreviewChanges(ResourceFile file, string resourceName) {
            Project project = null;
            try { project = this.StringToExtract.Parent.Document.ProjectItem.ContainingProject; } catch { }
            string reference = this.actionObject.GetResourceReference(file, resourceName, project);
            reference = this.StringToExtract.GetShortestReference(reference, this.StringToExtract.GetImportedNamespaces());
            return reference;
        }

        /// <summary>
        /// Extracts the string to a resource file and replaces it with a reference to the new entry
        /// </summary>
        /// <param name="file"></param>
        /// <param name="resourceName"></param>
        public void ExtractStringToResource(ResourceFile file, string resourceName) {
            UndoContext undo = StringToExtract.Parent.Document.DTE.UndoContext;
            undo.Open(Strings.ExtractResourceUndoContextName);
            try {
                Project project = null;
                try { project = this.StringToExtract.Parent.Document.ProjectItem.ContainingProject; } catch { }
                string reference = this.actionObject.GetResourceReference(file, resourceName, project);
                reference = this.StringToExtract.GetShortestReference(reference, this.StringToExtract.GetImportedNamespaces());
                this.StringToExtract.Replace(reference);
                undo.Close();
            } catch {
                undo.SetAborted();
                throw;
            }
        }

        #region Static Methods

        /// <summary>
        /// A cache of action items used before, key is a ProjectItem instance since a ProjectItem will always
        /// have the same action object
        /// </summary>
        private static Dictionary<ProjectItem, IExtractResourceAction> actionObjectCache = new Dictionary<ProjectItem, IExtractResourceAction>();

        /// <summary>
        /// List of available implementation in the current assembly.
        /// </summary>
        private static List<IExtractResourceAction> availableActionObjectList;

        /// <summary>
        /// Gets the correct ExtractResource action implementation for the provided string instance
        /// </summary>
        /// <param name="instance">String instance to refactor</param>
        /// <returns>an IExtractResourceAction implementation. null if no suitable action is found.</returns>
        public static IExtractResourceAction GetExtractResourceActionObject(BaseHardCodedString instance) {
            if (actionObjectCache.ContainsKey(instance.Parent)) return actionObjectCache[instance.Parent];
            if (availableActionObjectList == null) {
                availableActionObjectList = new List<IExtractResourceAction>();
                Assembly currentAssembly = Assembly.GetExecutingAssembly();
                Type[] types = currentAssembly.GetTypes();
                foreach (Type type in types) {
                    if (!type.IsAbstract && (typeof(IExtractResourceAction)).IsAssignableFrom(type)) {
                        try {
                            IExtractResourceAction currentAction = System.Activator.CreateInstance(type) as IExtractResourceAction;
                            if (currentAction != null) {
                                availableActionObjectList.Add(currentAction);
                            }
                        } catch (TargetInvocationException e) {
                            System.Diagnostics.Trace.TraceError(e.ToString());
                        } catch (ArgumentException e) {
                            System.Diagnostics.Trace.TraceError(e.ToString());
                        } catch (MissingMethodException e) {
                            System.Diagnostics.Trace.TraceError(e.ToString());
                        }
                    }
                }
            }
            IExtractResourceAction action = null;
            foreach (IExtractResourceAction currentAction in availableActionObjectList) {
                if (!currentAction.QuerySupportForProject(instance.Parent)) {
                    continue;
                }
                if (action == null || currentAction.Priority > action.Priority) {
                    action = currentAction;
                }
            }
            if (action != null) {
                if (actionObjectCache.Count > 30) actionObjectCache.Clear();
                actionObjectCache.Add(instance.Parent, action);
            }
            return action;
        }
        #endregion
    }



}
