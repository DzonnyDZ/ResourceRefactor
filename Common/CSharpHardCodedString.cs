/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections.ObjectModel;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>
    /// An implementation of BaseHardCodedString interface for C# files. Supports both normal strings and verbatim strings.
    /// </summary>
    public class CSharpHardCodedString : BaseHardCodedString {
        /// <summary>
        /// Cached value of the string
        /// </summary>
        private string value;

        /// <summary>
        /// Regex object for C# comments
        /// </summary>
        private static Regex commentRegexEngine = null;

        /// <summary>
        /// Constructor for hard coded strings in C#
        /// </summary>
        /// <param name="parent">Reference to code file containing the string</param>
        /// <param name="lineNumber">Line number</param>
        /// <param name="start">Starting index (including quotes)</param>
        /// <param name="end">Ending index (including quotes)</param>
        public CSharpHardCodedString(ProjectItem parent, int start, int end)
            :
            base(parent, start, end) {
        }

        /// <summary>
        /// Creates a new instance to use string checking functions.
        /// </summary>
        public CSharpHardCodedString()
            : base() {
        }

        #region BaseHardCodedString interface members

        /// <summary>
        /// String representation of the literal, this would be the value to be placed in to resource files.
        /// </summary>
        public override string Value {
            get {
                if (this.value == null) {
                    this.value = this.BeginEditPoint.GetText(this.TextLength);
                    if (this.value.StartsWith("@")) {
                        //Verbatim string
                        this.value = value.Substring(2, value.Length - 3).Replace("\"\"", "\"");
                    } else {
                        this.value = value.Substring(1, value.Length - 2);
                        value = Regex.Unescape(value);
                    }
                }
                return this.value;
            }
        }

        /// <summary>
        /// Creates another instance of CSharpHardCoded string with the provided arguments
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="lineNumber"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public override BaseHardCodedString CreateInstance(ProjectItem parent, int start, int end) {
            return new CSharpHardCodedString(parent, start, end);
        }

        #endregion

        #region BaseHardCodedString members

        protected override string StringRegExp {
            get {
                return Strings.RegexCSharpLiteral;
            }
        }

        protected override Regex CommentRegularExpression {
            get {
                if (commentRegexEngine == null) {
                    commentRegexEngine = new Regex(Strings.RegexCSharpComments, RegexOptions.Compiled);
                }
                return commentRegexEngine;
            }
        }

        public override System.Collections.ObjectModel.Collection<NamespaceImport> GetImportedNamespaces() {
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
                        FindUsingStatements(element, importedNamespaces);
                    }
                }
            } catch (System.Runtime.InteropServices.COMException) { } catch (NotImplementedException) { }
            return importedNamespaces;
        }

        /// <summary>
        /// Recurses through code elements to find all using statements and inserts them in to the provided collection
        /// </summary>
        /// <param name="element">Code element representing either namespace or import statement</param>
        /// <param name="namespaces">Collection to ad namespaces to</param>
        private void FindUsingStatements(CodeElement element, Collection<NamespaceImport> namespaces) {
            System.Text.RegularExpressions.Regex regExp = new Regex("using[ \\t]+((.)*);");
            if (element.Kind == vsCMElement.vsCMElementImportStmt) {
                string text = element.StartPoint.CreateEditPoint().GetText(element.EndPoint);
                Match m = regExp.Match(text);
                if (m.Groups.Count > 2) {
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
                    FindUsingStatements(children, namespaces);
                }
            }
        }
        #endregion

    }
}
