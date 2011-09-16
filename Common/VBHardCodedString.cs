/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections.ObjectModel;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common
{
    /// <summary>
    /// An implementation of BaseHardCodedString interface for VB files.
    /// </summary>
    public class VBHardCodedString : BaseHardCodedString
    {

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
        public VBHardCodedString(ProjectItem parent, int start, int end)
            :
            base(parent, start, end)
        {
        }

        /// <summary>
        /// Creates a new instance to use string checking functions.
        /// </summary>
        public VBHardCodedString() : base()
        {
        }

        #region BaseHardCodedString interface members

        public override string Value
        {
            get
            {
                if (this.value == null)
                {
                    this.value = this.BeginEditPoint.GetText(this.TextLength);
                    this.value = value.Substring(1, value.Length - 2).Replace("\"\"", "\"");
                }
                return this.value;
            }
        }

        /// <summary>
        /// Creates another instance of VBHardCoded string with the provided arguments
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="lineNumber"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public override BaseHardCodedString CreateInstance(ProjectItem parent, int start, int end)
        {
            return new VBHardCodedString(parent, start, end);
        }

        #endregion

        #region BaseHardCodedString members

        protected override string StringRegExp
        {
            get
            {
                return Strings.RegexVBLiteral;
            }
        }

        protected override Regex CommentRegularExpression
        {
            get
            {
                if (commentRegexEngine == null)
                {
                    commentRegexEngine = new Regex(Strings.RegexVBComments, RegexOptions.Compiled);
                }
                return commentRegexEngine;
            }
        }

        public override System.Collections.ObjectModel.Collection<NamespaceImport> GetImportedNamespaces()
        {
            Collection<NamespaceImport> importedNamespaces = new Collection<NamespaceImport>();
            if (this.Parent.FileCodeModel != null)
            {
                try
                {
                    CodeElement t = this.Parent.FileCodeModel.CodeElementFromPoint(
                                                    this.BeginEditPoint,
                                                    vsCMElement.vsCMElementNamespace);
                    if (t != null)
                    {
                        importedNamespaces.Add(new NamespaceImport(t.FullName, t.FullName));
                    }
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                }
                foreach (CodeElement element in this.Parent.FileCodeModel.CodeElements)
                {
                    FindUsingStatements(element, importedNamespaces);
                }
            }
            return importedNamespaces;
        }

        /// <summary>
        /// Recurses through code elements to find all using statements and inserts them in to the provided collection
        /// </summary>
        /// <param name="element"></param>
        /// <param name="namespaces"></param>
        private void FindUsingStatements(CodeElement element, Collection<NamespaceImport> namespaces)
        {
            System.Text.RegularExpressions.Regex regExp = new Regex(@"Imports[ \t](?<ns>[^,]+)(,(?<ns>[^,]+))*", RegexOptions.ExplicitCapture | RegexOptions.IgnoreCase);
            if (element.Kind == vsCMElement.vsCMElementImportStmt)
            {
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
            if (element.Kind == vsCMElement.vsCMElementNamespace)
            {
                foreach (CodeElement children in element.Children)
                {
                    FindUsingStatements(children, namespaces);
                }
            }
        }
        #endregion
    }
}
