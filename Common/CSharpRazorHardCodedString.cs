﻿using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Web.Configuration;
using System.Web.WebPages.Razor.Configuration;
using EnvDTE;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>An implementation of <see cref="BaseHardCodedString"/> interface for cshtml (Razor) files.</summary>
    public class CSharpRazorHardCodedString : CSharpHardCodedString {
        /// <summary>CTor - creates a new instance of the <see cref="CSharpRazorHardCodedString"/> class</summary>
        public CSharpRazorHardCodedString() { }

        /// <summary>Constructor for hard coded strings in C# Razor files</summary>
        /// <param name="parent">Reference to code file containing the string</param>
        /// <param name="lineNumber">Line number</param>
        /// <param name="start">Starting index (including quotes)</param>
        /// <param name="end">Ending index (including quotes)</param>
        public CSharpRazorHardCodedString(ProjectItem parent, int start, int end) : base(parent, start, end) { }

        /// <summary>Creates another instance of <see cref="CSharpRazorHardCodedString"/> string with the provided arguments</summary>
        /// <param name="parent">Current object item (file)</param>
        /// <param name="start">Starting offset of the string</param>
        /// <param name="end">End offset of the string</param>
        /// <returns>A new instance of <see cref="CSharpRazorHardCodedString"/>.</returns>
        public override BaseHardCodedString CreateInstance(ProjectItem parent, int start, int end) {
            CSharpRazorHardCodedString newInstance = new CSharpRazorHardCodedString(parent, start, end);
            newInstance.InitializeValue();
            return newInstance;
        }

        /// <summary>Gets the regular expression to identify strings</summary>
        protected override string StringRegExp {
            get {
                return String.Format("{0}|{1}", Strings.RegexInnerXML, base.StringRegExp);
            }
        }

        /// <summary>Cached value of the string</summary>
        private string value;
        /// <summary>Is Razor prefix required?</summary>
        private bool needsRazorPrefix = false;

        private void InitializeValue() {
            this.value = this.BeginEditPoint.GetText(this.TextLength);
            if (this.value.StartsWith("@")) {
                //Verbatim string
                this.value = value.Substring(2, value.Length - 3).Replace("\"\"", "\"");
            }
            else if (this.value.StartsWith("\"")) {
                // String in quotes.
                this.value = value.Substring(1, value.Length - 2);
                value = Regex.Unescape(value);
            }
            else {
                // String in html tags.
                this.value = value.Trim();
                this.needsRazorPrefix = true;
            }
        }
        /// <summary>Gets string representation of the literal, this would be the value to be placed in to resource files.</summary>
        public override string Value {
            get {
                if (this.value == null) {
                    InitializeValue();
                }
                return this.value;

            }
        }

        /// <summary>Returns a collection of namespaces imported in current file or in web.config files</summary>
        /// <returns>A collection of namespace imports for effective for hardcoded string location</returns>
        /// <remarks>This list will be used to determine the replacement string</remarks>
        public override Collection<NamespaceImport> GetImportedNamespaces() {
            var namespaces = base.GetImportedNamespaces();
            GetWebConfigNamespaces(namespaces);
            GetNamespacesFromFile(namespaces);

            return namespaces;
        }

        /// <summary>Returns a collection of namespaces imported in the CSHTML file itself</summary>
        /// <param name="namespaces">A collection to add namespaces imported in the CSHTML file to</param>
        private void GetNamespacesFromFile(Collection<NamespaceImport> namespaces) {
            TextDocument doc = (TextDocument)this.Parent.Document.Object("TextDocument");
            string contents = ExtensibilityMethods.GetDocumentText(doc);
            System.Text.RegularExpressions.Regex regExp = new Regex("@using[ \\t]+((.)*)");
            Match m = regExp.Match(contents);

            if (m.Groups.Count > 0) {
                if (m.Groups[1].Value.Contains("=")) {
                    var parts = m.Groups[1].Value.Split(new[] { '=' }, 2);
                    namespaces.Add(new NamespaceImport(m.Groups[1].Value, parts[0].Trim(), parts[1].Trim()));
                } else {
                    string ns = m.Groups[1].Value.Trim();
                    namespaces.Add(new NamespaceImport(ns, ns));
                }
            }
        }

        /// <summary>Returns a collection of namespaces imported in web.config files</summary>
        /// <param name="namespaces">A collection to add namespaces imported in web.config files to</param>
        private void GetWebConfigNamespaces(Collection<NamespaceImport> namespaces) {
            string currentPath = this.Parent.Document.Path; // "D:\\CarolesFiles\\Documents\\Visual Studio 2010\\Projects\\OdeToFood\\OdeToFood\\Views\\Home\\";
            string projectPath = this.Parent.ContainingProject.FullName;  //"D:\\CarolesFiles\\Documents\\Visual Studio 2010\\Projects\\OdeToFood\\OdeToFood\\OdeToFood.csproj";
            projectPath = Path.GetDirectoryName(projectPath);

            var configFileMap = new WebConfigurationFileMap();
            var virtualDirectories = configFileMap.VirtualDirectories;
            string directoryVirtualPath = null;

            while (!currentPath.Equals(projectPath, StringComparison.OrdinalIgnoreCase)) {
                currentPath = Path.GetDirectoryName(currentPath);  // Gets the path of the current path's parent
                string relativePath = currentPath.Substring(projectPath.Length);

                bool isAppRoot = currentPath.Equals(projectPath, StringComparison.OrdinalIgnoreCase);
                string virtualPath = relativePath.Replace('\\', '/');
                if (virtualPath.Length == 0) {
                    virtualPath = "/";
                }

                directoryVirtualPath = directoryVirtualPath ?? virtualPath;

                virtualDirectories.Add(virtualPath, new VirtualDirectoryMapping(currentPath, isAppRoot: isAppRoot));
            }

            var config = WebConfigurationManager.OpenMappedWebConfiguration(configFileMap, directoryVirtualPath);

            // We use dynamic here because we could be dealing both with a 1.0 or a 2.0 RazorPagesSection, which
            // are not type compatible.
            dynamic section = config.GetSection(RazorPagesSection.SectionName);
            if (section != null) {
                foreach (NamespaceInfo n in section.Namespaces) {
                    Debug.WriteLine(n.Namespace);
                    namespaces.Add(new NamespaceImport(n.Namespace, n.Namespace));
                }

            }
        }     

        /// <summary>Shortens a full namespace reference by looking at a list of namespaces that are imported in the code</summary>
        /// <param name="reference">Reference to shorten</param>
        /// <param name="namespaces">Collection of namespaces imported in the file</param>
        /// <returns>Shortest form the of the reference valid for the file</returns>
        /// <remarks>This implementation is same for C# and VB</remarks>
        public override string GetShortestReference(string reference, Collection<NamespaceImport> namespaces) {
            string refstr = base.GetShortestReference(reference, namespaces);

            if (this.needsRazorPrefix)
                refstr = refstr.Insert(0, "@");
            return refstr;
        }

    }
}
