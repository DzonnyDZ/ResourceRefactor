/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections.ObjectModel;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>An implementation of <see cref="BaseHardCodedString"/> for VB files.</summary>
    public class VBHardCodedString : BaseHardCodedString {

        /// <summary>Cached value of the string</summary>
        private string value;

        /// <summary>Regex object for VB comments</summary>
        private static Regex commentRegexEngine = null;

        /// <summary>Constructor for hard coded strings in VB</summary>
        /// <param name="parent">Reference to code file containing the string</param>
        /// <param name="lineNumber">Line number</param>
        /// <param name="start">Starting index (including quotes)</param>
        /// <param name="end">Ending index (including quotes)</param>
        public VBHardCodedString(ProjectItem parent, int start, int end) : base(parent, start, end) { }

        /// <summary>Default CTor - Creates a new instance to use string checking functions.</summary>
        public VBHardCodedString() { }

        #region BaseHardCodedString members

        /// <summary>When overrden in derived class gets actual value of the string (without quotes and with special characters replaced)</summary>
        public override string Value {
            get {
                if (this.value == null) {
                    this.value = this.BeginEditPoint.GetText(this.TextLength);
                    this.value = value.Substring(1, value.Length - 2).Replace("\"\"", "\"");
                }
                return this.value;
            }
        }

        /// <summary>Creates another instance of VBHardCoded string with the provided arguments</summary>
        /// <param name="parent">Current object item (file)</param>
        /// <param name="start">Starting offset of the string</param>
        /// <param name="end">End offset of the string</param>
        /// <returns>A new instance of <see cref="VBHardCodedString"/>.</returns>
        public override BaseHardCodedString CreateInstance(ProjectItem parent, int start, int end) {
            return new VBHardCodedString(parent, start, end);
        }

        #endregion

        #region BaseHardCodedString members

        /// <summary>Gets the regular expression to identify strings</summary>
        protected override string StringRegExp { get { return Strings.RegexVBLiteral; } }

        /// <summary>Gets the regular expression object sting to identify comments</summary>
        protected override Regex CommentRegularExpression {
            get {
                if (commentRegexEngine == null) {
                    commentRegexEngine = new Regex(Strings.RegexVBComments, RegexOptions.Compiled);
                }
                return commentRegexEngine;
            }
        }

        /// <summary>Returns a collection of namespaces imported in the files ('Imports' keyword)</summary>
        /// <returns>A collection of namespace imports for effective for hardcoded string location</returns>
        /// <remarks>This list will be used to determine the replacement string</remarks>
        public override Collection<NamespaceImport> GetImportedNamespaces() {
            Collection<NamespaceImport> importedNamespaces = new Collection<NamespaceImport>();
            try {
                if (this.Parent.FileCodeModel != null) {
                    try {
                        CodeElement t = this.Parent.FileCodeModel.CodeElementFromPoint(
                                                        this.BeginEditPoint,
                                                        vsCMElement.vsCMElementNamespace);
                        if (t != null) {
                            importedNamespaces.Add(new NamespaceImport(t.FullName, t.FullName));
                        }
                    } catch (System.Runtime.InteropServices.COMException) {
                    }
                    foreach (CodeElement element in this.Parent.FileCodeModel.CodeElements) {
                        FindImportsStatements(element, importedNamespaces);
                    }
                }
            } catch (System.Runtime.InteropServices.COMException) { } catch (NotImplementedException) { }
            return importedNamespaces;
        }

        /// <summary>Recurses through code elements to find all Imports statements and inserts them in to the provided collection</summary>
        /// <param name="element">Current code element to examine</param>
        /// <param name="namespaces">A collection where to add any discovered namespaces</param>
        private void FindImportsStatements(CodeElement element, Collection<NamespaceImport> namespaces) {
            System.Text.RegularExpressions.Regex regExp = new Regex(@"Imports[ \t](?<ns>[^,]+)(,(?<ns>[^,]+))*", RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase);
            if (element.Kind == vsCMElement.vsCMElementImportStmt) {
                string text = element.StartPoint.CreateEditPoint().GetText(element.EndPoint);
                Match m = regExp.Match(text);
                if (m.Success && m.Groups.Count > 1) {
                    if (m.Groups[1].Value.Contains("=")) {
                        var parts = m.Groups[1].Value.Split(new[] { '=' }, 2);
                        namespaces.Add(new NamespaceImport(m.Groups[1].Value, parts[0].Trim(), parts[1].Trim()));
                    } else
                        namespaces.Add(new NamespaceImport(m.Groups[1].Value, m.Groups[1].Value));
                }
            }

            // We don't need to recurse in to other types as they won't contain using statements
            if (element.Kind == vsCMElement.vsCMElementNamespace) {
                foreach (CodeElement children in element.Children) {
                    FindImportsStatements(children, namespaces);
                }
            }
        }
        #endregion
    }
}