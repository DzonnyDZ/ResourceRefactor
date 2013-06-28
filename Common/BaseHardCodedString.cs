using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics.CodeAnalysis;
using System.Collections.ObjectModel;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>Represents a hard coded string in a code file. </summary>
    /// <remarks><note type="inheritinfo">Implementations are responsible for checking for hard coded strings, previewing changes and performing replacements.</note></remarks>
    public abstract class BaseHardCodedString {
        #region Private Members
        private ProjectItem parent;
        private Regex literalRegex;
        private EditPoint beginEditPoint;
        private EditPoint endEditPoint;
        private int textLength;
        #endregion

        /// <summary>Gets <see cref="Regex"/> parser for literals</summary>
        /// <remarks><see cref="Regex"/> is obtained form <see cref="StringRegExp"/> and compiled once.</remarks>
        private Regex LiteralRegex {
            get {
                if (literalRegex == null)
                    this.literalRegex = new Regex(this.StringRegExp, RegexOptions.Compiled);
                return literalRegex;
            }
        }

        #region Protected Properties

        /// <summary>Gets edit point to be used to replace the string or read the lines (start point)</summary>
        protected EditPoint BeginEditPoint { get { return beginEditPoint; } }

        /// <summary>Gets edit point to be used to replace the string or read the lines (end point)</summary>
        protected EditPoint EndEditPoint { get { return endEditPoint; } }

        /// <summary>Gets length of the string including quotes</summary>
        protected int TextLength { get { return textLength; } }

        /// <summary>When overriden in derived class gets the regular expression to identify strings</summary>
        protected abstract string StringRegExp { get; }

        /// <summary>When overriden in derived class gets the regular expression object sting to identify comments</summary>
        protected abstract Regex CommentRegularExpression { get; }

        #endregion

        #region Public Properties
        /// <summary>Gets reference to the file containing the string</summary>
        public EnvDTE.ProjectItem Parent { get { return parent; } }

        /// <summary>Gets aobsolute position (in character) of edit point to be used to replace the string or read the lines (start point)</summary>
        public int AbsoluteStartIndex { get { return BeginEditPoint.AbsoluteCharOffset - 1; } }

        /// <summary>Gets aobsolute position (in character) of edit point to be used to replace the string or read the lines (end point)</summary>
        public int AbsoluteEndIndex { get { return this.endEditPoint.AbsoluteCharOffset - 1; } }

        /// <summary>Gets starting index of hard coded literal (including possible quotes)</summary>
        public int StartIndex { get { return this.beginEditPoint.LineCharOffset - 1; } }

        /// <summary>Gets 0-based index of the starting line of the string</summary>
        public int StartingLine { get { return this.beginEditPoint.Line - 1; } }

        /// <summary>Gets 0-based index of the last line of the string</summary>
        public int EndingLine { get { return this.endEditPoint.Line - 1; } }

        /// <summary>Gets ending index of hard coded literal</summary>
        public int EndIndex { get { return this.endEditPoint.LineCharOffset - 1; } }

        /// <summary>Gets the raw value of the hard coded string as it is represented in the code file.</summary>
        public string RawValue { get { return this.beginEditPoint.GetText(this.TextLength); } }

        /// <summary>Gets the project type for the project that contains the item</summary>
        public ProjectType ProjectType {
            get { return ExtensibilityMethods.GetProjectType(this.Parent.ContainingProject); }
        }

        /// <summary>When overrden in derived class gets actual value of the string (without quotes and with special characters replaced)</summary>
        public abstract string Value { get; }
        #endregion

        #region CTors

        /// <summary>CTor - creates a new generic hard coded string pointer.</summary>
        /// <param name="parent">Reference to code file containing the string</param>
        /// <param name="lineNumber">Line number</param>
        /// <param name="start">Starting index (including quotes, with 0 indexing)</param>
        /// <param name="end">Ending index (including quotes)</param>
        protected BaseHardCodedString(ProjectItem parent, int start, int end)
            : this() {
            this.parent = parent;
            TextDocument doc = BaseHardCodedString.GetDocumentForItem(parent);
            this.beginEditPoint = doc.StartPoint.CreateEditPoint();
            if (start < Int32.MaxValue) start++;
            if (end < Int32.MaxValue) end++;
            this.beginEditPoint.MoveToAbsoluteOffset(start);
            this.endEditPoint = doc.StartPoint.CreateEditPoint();
            this.endEditPoint.MoveToAbsoluteOffset(end);
            this.textLength = end - start;
        }

        /// <summary>Default CTor - creates a new instance to use string checking functions.</summary>
        protected BaseHardCodedString() { }

        #endregion

        #region Instance Methods

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
        public virtual string Replace(string text) {
            this.parent.Document.Activate();
            if (!ExtensibilityMethods.CheckoutItem(this.Parent)) throw new FileCheckoutException(this.Parent.Name);
            if (this.parent.Document.ReadOnly) throw new FileReadOnlyException(this.Parent.Name);

            this.beginEditPoint.Delete(this.textLength);
            this.beginEditPoint.Insert(text);
            return this.beginEditPoint.GetLines(this.beginEditPoint.Line, this.beginEditPoint.Line + 1);
        }

        /// <summary>Checks if selection in <see cref="TextDocument"/> is inside a hard coded string and return a <see cref="MatchResult"/> pointing to string.</summary>
        /// <param name="document">A <see cref="TextDocument"/> to look in part of</param>
        /// <param name="line">Current line</param>
        /// <param name="selectionStart">Starting index of the selection</param>
        /// <param name="selectionEnd">Ending index of the selection</param>
        /// <returns>A <see cref="MatchResult"/> object, result is false if a string could not be located.</returns>
        public MatchResult CheckForHardCodedString(TextDocument document, int selectionStart, int selectionEnd) {
            return CheckForHardCodedString(ExtensibilityMethods.GetDocumentText(document).Replace("\r\n", "\n"), selectionStart, selectionEnd);
        }

        /// <summary>Checks if selection in <see cref="string"/> is inside a hard coded string and return a <see cref="MatchResult"/> pointing to string.</summary>
        /// <param name="text">A string too look inside</param>
        /// <param name="selectionStart">Starting index of the selection</param>
        /// <param name="selectionEnd">Ending index of the selection</param>
        /// <returns>A <see cref="MatchResult"/> object, result is false if a string could not be located.</returns>
        public MatchResult CheckForHardCodedString(string text, int selectionStart, int selectionEnd) {
            MatchResult result = new MatchResult();
            MatchCollection matches = this.LiteralRegex.Matches(text);
            foreach (Match match in matches) {
                if (match.Index <= selectionStart &&
                    match.Index + match.Length >= selectionEnd) {
                    result.Result = true;
                    result.StartIndex = match.Index;
                    result.EndIndex = match.Index + match.Length;
                    break;
                }
            }
            return result;
        }

        /// <summary>When overriden in derived class creates another instance of hardcoded string with the same type as the callee object.</summary>
        /// <param name="parent">Current object item (file)</param>
        /// <param name="start">Starting offset of the string</param>
        /// <param name="end">End offset of the string</param>
        /// <returns>A new instance of <see cref="BaseHardCodedString"/>-derived object.</returns>
        public abstract BaseHardCodedString CreateInstance(ProjectItem parent, int start, int end);

        /// <summary>When overriden in derived class returns a collection of namespaces imported in the files ('using' keyword in C#, or 'Imports' in VB.Net)</summary>
        /// <returns>A collection of namespace imports for effective for hardcoded string location</returns>
        /// <remarks>This list will be used to determine the replacement string</remarks>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")]
        public abstract Collection<NamespaceImport> GetImportedNamespaces();

        /// <summary>Returns the position of all comments in the given document using regular expressions</summary>
        /// <param name="item">Project item representing the document</param>
        /// <returns><see cref="Match"/> objects representing positions of all comments in the given documen</returns>
        public MatchCollection FindCommentsInDocument(ProjectItem item) {
            return this.CommentRegularExpression.Matches(ExtensibilityMethods.GetDocumentText(GetDocumentForItem(item)).Replace("\r\n", "\n"));
        }

        /// <summary>Shortens a full namespace reference by looking at a list of namespaces that are imported in the code</summary>
        /// <param name="reference">Reference to shorten</param>
        /// <param name="namespaces">Collection of namespaces imported in the file</param>
        /// <returns>Shortest form the of the reference valid for the file</returns>
        /// <remarks>This implementation is same for C# and VB</remarks>
        public virtual string GetShortestReference(string reference, Collection<NamespaceImport> namespaces) {
            NamespaceImport longestMatch = null;
            // Retrieve the property part of the reference
            string property = reference;
            string[] parts = reference.Split('.');
            property = parts[parts.Length - 1];
            reference = reference.Substring(0, reference.Length - property.Length - (parts.Length == 0 ? 0 : 1));
            foreach (var ns in namespaces) {

                if (!reference.Equals(ns.NamespaceName) && reference.StartsWith(ns + ".") &&
                     (longestMatch == null || ns.NamespaceName.Length > longestMatch.NamespaceName.Length)) {
                    longestMatch = ns;
                }
            }
            if (longestMatch != null)
                reference = reference.Remove(0, longestMatch.NamespaceName.Length);
            if (reference.StartsWith(".")) {
                reference = reference.Remove(0, 1);
            }
            if (longestMatch != null && longestMatch.Alias != null)
                reference = longestMatch.Alias + "." + longestMatch.NamespaceName;
            reference += "." + property;
            return reference;
        }

        #endregion

        #region Static Methods

        /// <summary>
        /// Returns an instance of BaseHardCodedString, based the type of the current document.
        /// </summary>
        /// <param name="currentDocument">The document that contains an hard-coded string.</param>
        /// <returns>A hard coded string that is file type aware.</returns>
        public static BaseHardCodedString GetHardCodedString(Document currentDocument) {
            BaseHardCodedString stringInstance = null;

            // Create the hard coded string instance
            switch (currentDocument.Language) {
                case "CSharp":
                    stringInstance = new CSharpHardCodedString();
                    break;
                case "Basic":
                    stringInstance = new VBHardCodedString();
                    break;
                case "XAML":
                    stringInstance = new XamlHardCodedString();
                    break;
                case "HTML":
                    if (currentDocument.Name.EndsWith(".cshtml", StringComparison.CurrentCultureIgnoreCase)) {
                        stringInstance = new CSharpRazorHardCodedString();
                    }
                    else if (currentDocument.Name.EndsWith(".vbhtml", StringComparison.CurrentCultureIgnoreCase)) {
                        stringInstance = new VBRazorHardCodedString();
                    }
                    else if (currentDocument.Name.EndsWith(".aspx", StringComparison.CurrentCultureIgnoreCase) || currentDocument.Name.EndsWith(".ascx", StringComparison.CurrentCultureIgnoreCase) || currentDocument.Name.EndsWith(".master", StringComparison.CurrentCultureIgnoreCase)) {
                        stringInstance = new AspxHardCodedString();
                    }
                    break;
            }
            return stringInstance;

        }


        /// <summary>Finds all instances of a text in the document object for the provided item.</summary>
        /// <param name="item">Project item to search texts</param>
        /// <param name="text">Text to look for</param>
        /// <returns>A collection of BaseHardCodedString objects that references to text</returns>
        public static ReadOnlyCollection<BaseHardCodedString> FindAllInstancesInDocument(ProjectItem item, string text) {
            TextDocument doc = GetDocumentForItem(item);
            Collection<BaseHardCodedString> instances = new Collection<BaseHardCodedString>();
            EditPoint start = doc.StartPoint.CreateEditPoint();
            start.MoveToAbsoluteOffset(1);
            EditPoint end = null;
            // Unused object for FindPattern method.
            TextRanges ranges = null;
            System.Collections.IEnumerator comments = null;
            Match currentMatch = null;

            BaseHardCodedString instance = BaseHardCodedString.GetHardCodedString(item.Document);

            while (start.FindPattern(text, (int)(vsFindOptions.vsFindOptionsMatchCase), ref end, ref ranges)) {
                if (instance != null) {
                    if (comments == null) comments = instance.FindCommentsInDocument(item).GetEnumerator();
                    bool inComment = false;
                    while (currentMatch != null || (comments.MoveNext())) {
                        currentMatch = (Match)(comments.Current);
                        // If this match appears earlier then current text skip
                        if (currentMatch.Index + currentMatch.Length <= (start.AbsoluteCharOffset - 1)) {
                            currentMatch = null;
                        }
                        // If this comment is later then current text stop processing comments
                        else if (currentMatch.Index >= (end.AbsoluteCharOffset - 1)) {
                            break;
                        }
                        // At this point current text must be part of a comment block
                        else {
                            inComment = true;
                            break;
                        }
                    }
                    if (!inComment) {
                        instance = instance.CreateInstance(item, start.AbsoluteCharOffset - 1, end.AbsoluteCharOffset - 1);
                        instances.Add(instance);
                    }
                }
                start = end;
            }
            return new ReadOnlyCollection<BaseHardCodedString>(instances);
        }

        /// <summary>Finds all instances of a text in the code files contained by the provided project</summary>
        /// <param name="project">Project to search</param>
        /// <param name="text">Text to look for</param>
        /// <returns>A collection of BaseHardCodedString implementations</returns>
        public static ReadOnlyCollection<BaseHardCodedString> FindAllInstancesInProject(Project project, string text) {
            Collection<BaseHardCodedString> instances = new Collection<BaseHardCodedString>();
            CodeFileCollection codeFiles = new CodeFileCollection(project);
            foreach (ProjectItem item in codeFiles) {
                foreach (BaseHardCodedString instance in FindAllInstancesInDocument(item, text)) {
                    instances.Add(instance);
                }
            }
            return new ReadOnlyCollection<BaseHardCodedString>(instances);
        }

        /// <summary>Gets the <see cref="TextDocument"/> interface for the item provided</summary>
        /// <param name="item">The item to get <see cref="TextDocument"/> for</param>
        /// <returns>A <see cref="TextDocument"/> object obtained for <paramref name="item"/></returns>
        /// <exception cref="ArgumentNullException"><paramref name="item"/> is null</exception>
        public static TextDocument GetDocumentForItem(ProjectItem item) {
            if (item == null) throw new ArgumentNullException("item");

            if (item.Document == null) {
                item.Open(null);
            }
            return ((EnvDTE.TextDocument)item.Document.Object(null));
        }

        #endregion
    }

    /// <summary>Match result structure for hard coded string queries</summary>
    public struct MatchResult {
        /// <summary>Gets 0-based starting index of the matched part</summary>
        public int StartIndex { get; set; }

        /// <summary>Gets 0-based ending index of the matched part</summary>
        public int EndIndex { get; set; }

        /// <summary>Result of the match, if false no match was found and <see cref="StartIndex"/>, <see cref="EndIndex"/> properties will be 0</summary>
        public bool Result { get; set; }

        /// <summary>Indicates whether this instance and a specified object are equal.</summary>
        /// <returns>true if <paramref name="obj"/> and this instance are the same type and represent the same value; otherwise, false.</returns>
        /// <param name="obj">Another object to compare to. </param>
        /// <filterpriority>2</filterpriority>
        public override bool Equals(object obj) {
            if (obj is MatchResult) {
                MatchResult other = (MatchResult)obj;
                return this.Result == other.Result && this.StartIndex == other.StartIndex && this.EndIndex == other.EndIndex;
            }
            return base.Equals(obj);
        }

        /// <summary>Test two <see cref="MatchResult"/> objects for equality</summary>
        /// <param name="a">A <see cref="MatchResult"/></param>
        /// <param name="b">A <see cref="MatchResult"/></param>
        /// <returns>True if <paramref name="a"/> equals <paramref name="b"/>; false otherwise</returns>
        public static bool operator ==(MatchResult a, MatchResult b) {
            return a.Equals(b);
        }

        /// <summary>Test two <see cref="MatchResult"/> objects for inequality</summary>
        /// <param name="a">A <see cref="MatchResult"/></param>
        /// <param name="b">A <see cref="MatchResult"/></param>
        /// <returns>False if <paramref name="a"/> equals <paramref name="b"/>; true otherwise</returns>
        public static bool operator !=(MatchResult obj1, MatchResult obj2) {
            return !obj1.Equals(obj2);
        }

        /// <summary>Returns the hash code for this instance.</summary>
        /// <returns>A 32-bit signed integer that is the hash code for this instance.</returns>
        /// <filterpriority>2</filterpriority>
        public override int GetHashCode() {
            return StartIndex + EndIndex + (Result ? 0 : 1);
        }
    }
}
