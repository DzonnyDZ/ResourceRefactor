/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Text.RegularExpressions;
using System.IO;
using System.Collections.ObjectModel;
using System.Xml.Linq;
using System.Xml;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>An implementation of <see cref="BaseHardCodedString"/> for XAML files.</summary>
    public class XamlHardCodedString : BaseHardCodedString {

        /// <summary>Cached value of the string</summary>
        private string value;

        /// <summary>Regex object for XAML (XML) comments</summary>
        private static Regex commentRegexEngine = null;

        /// <summary>Constructor for hard coded strings in XAML</summary>
        /// <param name="parent">Reference to code file containing the string</param>
        /// <param name="lineNumber">Line number</param>
        /// <param name="start">Starting index (including quotes)</param>
        /// <param name="end">Ending index (including quotes)</param>
        public XamlHardCodedString(ProjectItem parent, int start, int end) : base(parent, start, end) { }

        /// <summary>Creates a new instance to use string checking functions.</summary>
        public XamlHardCodedString() { }

        #region BaseHardCodedString members

        /// <summary>When overrden in derived class gets actual value of the string (without quotes and with special characters replaced)</summary>
        public override string Value {
            get {
                if (this.value == null) {
                    this.value = this.BeginEditPoint.GetText(this.TextLength);
                    if (this.value.StartsWith("\"") && this.value.EndsWith("\"")) this.value = this.value.Substring(1, this.value.Length - 2);
                    else if (this.value.StartsWith("'") && this.value.EndsWith("'")) this.value = this.value.Substring(1, this.value.Length - 2);
                    XmlDocument doc = new XmlDocument();//XML unescape
                    XmlElement el = doc.CreateElement("aaa");
                    el.InnerXml = value;

                    this.value = el.InnerText;
                }
                return this.value;
            }
        }

        /// <summary>Creates another instance of <see cref="XamlHardCodedString"/> with the provided arguments</summary>
        /// <param name="parent">Current object item (file)</param>
        /// <param name="start">Starting offset of the string</param>
        /// <param name="end">End offset of the string</param>
        /// <returns>A new instance of <see cref="XamlHardCodedString"/>.</returns>
        public override BaseHardCodedString CreateInstance(ProjectItem parent, int start, int end) {
            return new XamlHardCodedString(parent, start, end);
        }

        #endregion

        #region BaseHardCodedString members

        /// <summary>Gets the regular expression to identify strings</summary>
        protected override string StringRegExp { get { return Strings.RegexXamlLiteral; } }

        /// <summary>Gets the regular expression object sting to identify comments</summary>
        protected override Regex CommentRegularExpression {
            get {
                if (commentRegexEngine == null) {
                    commentRegexEngine = new Regex(Strings.RegexXamlComments, RegexOptions.Compiled);
                }
                return commentRegexEngine;
            }
        }

        /// <summary>Returns a collection of namespaces imported in the files</summary>
        /// <returns>A collection of namespace imports for effective for hardcoded string location</returns>
        /// <remarks>This list will be used to determine the replacement string</remarks>
        public override Collection<NamespaceImport> GetImportedNamespaces() {
            Collection<NamespaceImport> importedNamespaces = new Collection<NamespaceImport>();
            TextSelection selection = (TextSelection)Parent.Document.Selection;
            var xamlDoc = selection.Parent;
            var start = xamlDoc.CreateEditPoint(xamlDoc.StartPoint);
            var end = xamlDoc.CreateEditPoint(xamlDoc.EndPoint);
            string allText = start.GetText(end);
            using (TextReader tr = new StringReader(allText))
            using (XmlReader xr = XmlReader.Create(tr)) {
                while (xr.Read()) {
                    if (xr.NodeType == XmlNodeType.Element) {//First element
                        if (xr.MoveToFirstAttribute()) {
                            do {
                                if (xr.Prefix == "xmlns") {
                                    importedNamespaces.Add(new NamespaceImport(string.Format("xmlns:{0}=\"{1}\"", xr.LocalName, xr.Value), xr.Value, xr.LocalName));
                                } else if (string.IsNullOrEmpty(xr.Prefix) && xr.LocalName == "xmlns") {
                                    importedNamespaces.Add(new NamespaceImport(string.Format("xmlns=\"{0}\"", xr.Value), xr.Value));
                                }
                            } while (xr.MoveToNextAttribute());
                        }
                        break;
                    }
                }
            }
            return importedNamespaces;
        }

        /// <summary>Shortens a full namespace reference by looking at a list of namespaces that are imported in the code</summary>
        /// <param name="reference">Reference to shorten</param>
        /// <param name="namespaces">Collection of namespaces imported in the file</param>
        /// <returns>Shortest form the of the reference valid for the file.</returns>
        public override string GetShortestReference(string reference, Collection<NamespaceImport> namespaces) {
            // Retrieve the property part of the reference
            string property = reference;
            string[] parts = reference.Split(new[] { ':' }, 2);
            string typeAndProperty = parts[1];
            string @namespace = parts[0];
            foreach (var ns in namespaces) {
                if (ns.NamespaceName == string.Format("clr-namespace:{0}", @namespace)) {
                    if (ns.Alias != null) return string.Format("{{x:Static {0}:{1}}}", ns.Alias, typeAndProperty);
                    else return string.Format("{{x:Static {1}}}", typeAndProperty);
                }
            }

            TextSelection selection = (TextSelection)Parent.Document.Selection;
            var xamlDoc = selection.Parent;
            var start = xamlDoc.CreateEditPoint(xamlDoc.StartPoint);
            var end = xamlDoc.CreateEditPoint(xamlDoc.EndPoint);
            string allText = start.GetText(end);
            using (TextReader tr = new StringReader(allText))
            using (XmlReader xr = XmlReader.Create(tr)) {
                while (xr.Read()) {
                    if (xr.NodeType == XmlNodeType.Element) {//First element
                        if (xr.MoveToFirstAttribute()) {
                            do { } while (xr.MoveToNextAttribute());
                            xr.ReadAttributeValue();
                        }
                        IXmlLineInfo linfo = (IXmlLineInfo)xr;
                        this.Parent.Document.Activate();
                        if (!ExtensibilityMethods.CheckoutItem(this.Parent))
                            throw new FileCheckoutException(this.Parent.Name);
                        if (this.Parent.Document.ReadOnly)
                            throw new FileReadOnlyException(this.Parent.Name);
                        var ep = xamlDoc.CreateEditPoint();
                        ep.MoveToLineAndOffset(linfo.LineNumber, linfo.LinePosition);
                        ep.MoveToAbsoluteOffset(ep.AbsoluteCharOffset + xr.Value.Length + 1);
                        ep.ReplaceText(ep, string.Format(" xmlns:{0}=\"clr-namespace:{0}\"", @namespace), (int)vsEPReplaceTextOptions.vsEPReplaceTextTabsSpaces);
                        return string.Format("{{x:Static {0}}}", reference);
                    }
                }
            }
            throw new InvalidOperationException("Neither namespace import (xmlns) nor position to insert it found");
        }
        #endregion

        /// <summary>Replaces the string with the specified text.</summary>
        /// <param name="text">Text to replace the current string</param>
        /// <returns>Full line where the text was replaced</returns>
        /// <remarks>
        /// Behaviour of BaseHardCodedString instance will be undetermined after this method is called, so object should not be used afterwards.
        /// <para>
        /// Changes should not be saved to the document immediatly.
        /// Once this method is called and returns successfully the object behaviour is allowed to become undefined.
        /// </para>
        /// </remarks>
        /// <exception cref="FileCheckoutException"><see cref="Parent"/> is under souzrce control and failed to check it out for editing</exception>
        /// <exception cref="FileReadOnlyException"><see cref="Parent"/> is read only</exception>
        override public string Replace(string text) {
            this.Parent.Document.Activate();
            if (!ExtensibilityMethods.CheckoutItem(this.Parent)) {
                throw new FileCheckoutException(this.Parent.Name);
            }
            if (this.Parent.Document.ReadOnly) {
                throw new FileReadOnlyException(this.Parent.Name);
            }
            this.BeginEditPoint.Delete(this.TextLength);
            this.BeginEditPoint.Insert("\"" + text + "\"");
            return this.BeginEditPoint.GetLines(this.BeginEditPoint.Line, this.BeginEditPoint.Line + 1);
        }
    }
}
