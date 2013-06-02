﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Configuration;
using System.Web.WebPages.Razor.Configuration;
using EnvDTE;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>An implementation of <see cref="BaseHardCodedString"/> for ASPX (ASCX and MASTER) files.</summary>
    public class AspxHardCodedString : BaseHardCodedString {
        /// <summary>Default CTor - creates a new instance of the <see cref="AspxHardCodedString"/> class</summary>
        public AspxHardCodedString()
            : base() { }

        /// <summary>Constructor for hard coded strings in ASPX markup files</summary>
        /// <param name="parent">Reference to code file containing the string</param>
        /// <param name="lineNumber">Line number</param>
        /// <param name="start">Starting index (including quotes)</param>
        /// <param name="end">Ending index (including quotes)</param>
        public AspxHardCodedString(ProjectItem parent, int start, int end)
            :
            base(parent, start, end) {
        }

        /// <summary>Creates another instance of CSharpHardCoded string with the provided arguments</summary>
        public override BaseHardCodedString CreateInstance(ProjectItem parent, int start, int end) {
            return new AspxHardCodedString(parent, start, end);
        }

        protected override string StringRegExp {
            get {
                return Strings.RegexAspxLiteral;
            }
        }

        protected override Regex CommentRegularExpression {
            get {
                return new Regex(Strings.RegexAspxComment);
            }
        }

        /// <summary>Cached value of the string</summary>
        private string value;

        public override string Value {
            get {
                if (this.value == null) {
                    this.value = this.BeginEditPoint.GetText(this.TextLength);
                    if (this.value.StartsWith("\"") || this.value.StartsWith("'")) {
                        this.value = HttpUtility.HtmlDecode(value.Substring(2, value.Length - 2));
                    } else {
                        this.value = HttpUtility.HtmlDecode(value);
                    }
                }
                return this.value;
            }
        }

        public override Collection<NamespaceImport> GetImportedNamespaces() {
            var namespaces = new Collection<NamespaceImport>();
            //GetWebConfigNamespaces(namespaces);
            //GetNamespacesFromFile(namespaces);
            return namespaces;
        }

        //private void GetNamespacesFromFile(Collection<NamespaceImport> namespaces) {
        //    TextDocument doc = (TextDocument)this.Parent.Document.Object("TextDocument");
        //    string contents = ExtensibilityMethods.GetDocumentText(doc);
        //    System.Text.RegularExpressions.Regex regExp = new Regex("@using[ \\t]+((.)*)");
        //    Match m = regExp.Match(contents);

        //    if (m.Groups.Count > 0) {
        //        if (m.Groups[1].Value.Contains("=")) {
        //            var parts = m.Groups[1].Value.Split(new[] { '=' }, 2);
        //            namespaces.Add(new NamespaceImport(m.Groups[1].Value, parts[0].Trim(), parts[1].Trim()));
        //        } else {
        //            string ns = m.Groups[1].Value.Trim();
        //            namespaces.Add(new NamespaceImport(ns, ns));
        //        }
        //    }
        //}

        //private void GetWebConfigNamespaces(Collection<NamespaceImport> namespaces) {
        //    string currentPath = this.Parent.Document.Path; // "D:\\CarolesFiles\\Documents\\Visual Studio 2010\\Projects\\OdeToFood\\OdeToFood\\Views\\Home\\";
        //    string projectPath = this.Parent.ContainingProject.FullName;  //"D:\\CarolesFiles\\Documents\\Visual Studio 2010\\Projects\\OdeToFood\\OdeToFood\\OdeToFood.csproj";
        //    projectPath = Path.GetDirectoryName(projectPath);

        //    var configFileMap = new WebConfigurationFileMap();
        //    var virtualDirectories = configFileMap.VirtualDirectories;
        //    string directoryVirtualPath = null;

        //    while (!currentPath.Equals(projectPath, StringComparison.OrdinalIgnoreCase)) {
        //        currentPath = Path.GetDirectoryName(currentPath);  // Gets the path of the current path's parent
        //        string relativePath = currentPath.Substring(projectPath.Length);

        //        bool isAppRoot = currentPath.Equals(projectPath, StringComparison.OrdinalIgnoreCase);
        //        string virtualPath = relativePath.Replace('\\', '/');
        //        if (virtualPath.Length == 0) {
        //            virtualPath = "/";
        //        }

        //        directoryVirtualPath = directoryVirtualPath ?? virtualPath;

        //        virtualDirectories.Add(virtualPath, new VirtualDirectoryMapping(currentPath, isAppRoot: isAppRoot));
        //    }

        //    var config = WebConfigurationManager.OpenMappedWebConfiguration(configFileMap, directoryVirtualPath);

        //    // We use dynamic here because we could be dealing both with a 1.0 or a 2.0 RazorPagesSection, which
        //    // are not type compatible.
        //    dynamic section = config.GetSection(RazorPagesSection.SectionName);
        //    if (section != null) {
        //        foreach (NamespaceInfo n in section.Namespaces) {
        //            Debug.WriteLine(n.Namespace);
        //            namespaces.Add(new NamespaceImport(n.Namespace, n.Namespace));
        //        }

        //    }
        //}


        //public override string GetShortestReference(string reference, Collection<NamespaceImport> namespaces) {
        //    string refstr = base.GetShortestReference(reference, namespaces);
        //    return refstr;
        //}

    }
}
