/// Copyright (c) Microsoft Corporation.  All rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using EnvDTE;
using System.Text.RegularExpressions;
using System.IO;
using System.Diagnostics.CodeAnalysis;
using System.Collections.ObjectModel;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common
{
    /// <summary>
    /// Represents a hard coded string in a code file. Implementations are responsible for
    /// checking for hard coded strings, previewing changes and performing replacements.
    /// </summary>
    public abstract class BaseHardCodedString
    {

        #region Private Members

        /// <summary>
        /// Reference to the file containing the string
        /// </summary>
        private ProjectItem parent;

        /// <summary>
        /// RegEx parser for literals
        /// </summary>
        private Regex literalRegex;

        /// <summary>
        /// Edit point to be used to replace the string or read the lines
        /// </summary>
        private EditPoint beginEditPoint;

        private EditPoint endEditPoint;

        /// <summary>
        /// Length of the string including quotes
        /// </summary>
        private int textLength;

        #endregion

        #region Protected Properties

        /// <summary>
        /// Edit point to be used to replace the string or read the lines
        /// </summary>
        protected EditPoint BeginEditPoint
        {
            get
            {
                return this.beginEditPoint;
            }
        }

        protected EditPoint EndEditPoint
        {
            get
            {
                return this.endEditPoint;
            }
        }

        /// <summary>
        /// Length of the string including quotes
        /// </summary>
        protected int TextLength
        {
            get
            {
                return textLength;
            }
        }

        /// <summary>
        /// Gets the regular expression to identify strings
        /// </summary>
        protected abstract string StringRegExp 
        {
            get;
        }

        /// <summary>
        /// Gets the regular expression object sting to identify comments
        /// </summary>
        protected abstract Regex CommentRegularExpression
        {
            get;
        }

        #endregion

        #region Public Properties
        /// <summary>
        /// Item containing the literal
        /// </summary>
        /// <seealso cref="BaseHardCodedString.Parent"/>
        public EnvDTE.ProjectItem Parent
        {
            get
            {
                return this.parent;
            }
        }

        public int AbsoluteStartIndex
        {
            get { return this.beginEditPoint.AbsoluteCharOffset - 1; }
        }

        public int AbsoluteEndIndex
        {
            get { return this.endEditPoint.AbsoluteCharOffset - 1; }
        }

        /// <summary>
        /// Starting index of hard coded literal (including " character)
        /// </summary>
        public int StartIndex
        {
            get
            {
                return this.beginEditPoint.LineCharOffset - 1;
            }
        }

        /// <summary>
        /// Returns the starting line of the string (0 indexed)
        /// </summary>
        public int StartingLine
        {
            get
            {
                return this.beginEditPoint.Line - 1;
            }
        }

        public int EndingLine
        {
            get { return this.endEditPoint.Line - 1; }
        }

        /// <summary>
        /// Ending index of hard coded literal
        /// </summary>
        public int EndIndex
        {
            get
            {
                return this.endEditPoint.LineCharOffset - 1;
            }
        }

        /// <summary>
        /// Gets the raw value of the hard coded string as it is represented in the code file.
        /// </summary>
        public string RawValue
        {
            get
            {
                return this.beginEditPoint.GetText(this.TextLength);
            }
        }

        /// <summary>
        /// Returns the project type for the project that contains the item
        /// </summary>
        /// <returns></returns>
        public ProjectType ProjectType
        {
            get
            {
                return ExtensibilityMethods.GetProjectType(this.Parent.ContainingProject);
            }
        }

        /// <summary>
        /// Actual value of the string (without quotes and special characters)
        /// </summary>
        public abstract string Value
        {
            get;
        }
        #endregion

        #region Constructors

        /// <summary>
        /// Creates a new generic hard coded string pointer.
        /// </summary>
        /// <param name="parent">Reference to code file containing the string</param>
        /// <param name="lineNumber">Line number</param>
        /// <param name="start">Starting index (including quotes, with 0 indexing)</param>
        /// <param name="end">Ending index (including quotes)</param>
        protected BaseHardCodedString(ProjectItem parent, int start, int end) : this()
        {
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

        /// <summary>
        /// Creates a new instance to use string checking functions.
        /// </summary>
        /// <remarks>Message is suppressed because this is the only way to perform the functionality.</remarks>
        [SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        protected BaseHardCodedString()
        {
            this.literalRegex = new Regex(this.StringRegExp, RegexOptions.Compiled);
        }

        #endregion

        #region Instance Methods 

        /// <summary>
        /// Replaces the string with the specified text.
        /// Changes should not be saved to the document immediatly. Once this method is called and returns successfully
        /// the object behaviour is allowed to become undefined.
        /// </summary>
        /// <param name="text">Text to replace the current string</param>
        /// <returns>Full line where the text was replaced</returns>
        /// <remarks>Behaviour of BaseHardCodedString instance will be undetermined after this method is called, so object should not be used afterwards</remarks>
        public virtual string Replace(string text)
        {
            this.parent.Document.Activate();
            if (!ExtensibilityMethods.CheckoutItem(this.Parent)) 
            {
                throw new FileCheckoutException(this.Parent.Name);
            }
            if (this.parent.Document.ReadOnly)
            {
                throw new FileReadOnlyException(this.Parent.Name);
            }
            this.beginEditPoint.Delete(this.textLength);
            this.beginEditPoint.Insert(text);
            return this.beginEditPoint.GetLines(this.beginEditPoint.Line, this.beginEditPoint.Line + 1);
        }

        /// <summary>
        /// Checks if selection is inside a hard coded string and return a MatchResult pointing to
        /// string.
        /// </summary>
        /// <param name="line">Current line</param>
        /// <param name="selectionStart">Starting index of the selection</param>
        /// <param name="selectionEnd">Ending index of the selection</param>
        /// <returns>a MatchResult object, result is false if a string could not be located.</returns>
        public MatchResult CheckForHardCodedString(TextDocument document, int selectionStart, int selectionEnd)
        {
            return CheckForHardCodedString(ExtensibilityMethods.GetDocumentText(document).Replace("\r\n","\n"), selectionStart, selectionEnd);
        }

        public MatchResult CheckForHardCodedString(string text, int selectionStart, int selectionEnd)
        {
            MatchResult result = new MatchResult();
            MatchCollection matches = this.literalRegex.Matches(text);
            foreach (Match match in matches)
            {
                if (match.Index <= selectionStart &&
                    match.Index + match.Length >= selectionEnd)
                {
                    result.Result = true;
                    result.StartIndex = match.Index;
                    result.EndIndex = match.Index + match.Length;
                    break;
                }
            }
            return result;
        }

        /// <summary>
        /// Creates another instance of hardcoded string with the same type as the callee object.
        /// </summary>
        /// <param name="parent"></param>
        /// <param name="start"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public abstract BaseHardCodedString CreateInstance(ProjectItem parent, int start, int end);

        /// <summary>
        /// Returns a collection of namespaces imported in the files ('using' keyword in C#, or 'Imports' in VB.Net)
        /// This list will be used to determine the replacement string
        /// </summary>
        /// <returns></returns>
        /// <remarks>CA1024: Since this is an expensive operation, it is implemented as a method</remarks>
        [SuppressMessage("Microsoft.Design", "CA1024:UsePropertiesWhereAppropriate")]
        public abstract Collection<NamespaceImport> GetImportedNamespaces();

        /// <summary>
        /// Returns the position of all comments in the given document using regular expressions
        /// </summary>
        /// <param name="item">Project item representing the document</param>
        /// <returns></returns>
        public MatchCollection FindCommentsInDocument(ProjectItem item)
        {
            return this.CommentRegularExpression.Matches(ExtensibilityMethods.GetDocumentText(GetDocumentForItem(item)).Replace("\r\n","\n"));
        }

        #endregion

        #region Static Methods

        /// <summary>
        /// Finds all instances of a text in the document object for the provided item.
        /// </summary>
        /// <param name="item">Project item to search texts</param>
        /// <param name="text">Text to look for</param>
        /// <returns>A collection of BaseHardCodedString objects that references to text</returns>
        public static ReadOnlyCollection<BaseHardCodedString> FindAllInstancesInDocument(ProjectItem item, string text)
        {
            TextDocument doc = GetDocumentForItem(item);
            Collection<BaseHardCodedString> instances = new Collection<BaseHardCodedString>();
            EditPoint start = doc.StartPoint.CreateEditPoint();
            start.MoveToAbsoluteOffset(1);
            EditPoint end = null;
            // Unused object for FindPattern method.
            TextRanges ranges = null;
            BaseHardCodedString instance = null;
            System.Collections.IEnumerator comments = null;
            Match currentMatch = null;
            switch (item.Document.Language)
            {
                case "CSharp":
                    instance = new Common.CSharpHardCodedString();
                    break;
                case "Basic":
                    instance = new Common.VBHardCodedString();
                    break;
                default:
                    break;
            }
            while (start.FindPattern(text, (int)(vsFindOptions.vsFindOptionsMatchCase), ref end, ref ranges))
            {
                if (instance != null)
                {
                    if (comments == null) comments = instance.FindCommentsInDocument(item).GetEnumerator();
                    bool inComment = false;
                    while (currentMatch != null || (comments.MoveNext()))
                    {
                        currentMatch = (Match)(comments.Current);
                        // If this match appears earlier then current text skip
                        if (currentMatch.Index + currentMatch.Length <= (start.AbsoluteCharOffset - 1))
                        {
                            currentMatch = null;
                        }
                        // If this comment is later then current text stop processing comments
                        else if (currentMatch.Index >= (end.AbsoluteCharOffset - 1))
                        {
                            break;
                        }
                        // At this point current text must be part of a comment block
                        else
                        {
                            inComment = true;
                            break;
                        }
                    }
                    if (!inComment) 
                    {
                        instance = instance.CreateInstance(item, start.AbsoluteCharOffset - 1, end.AbsoluteCharOffset - 1);
                        instances.Add(instance);
                    }
                }
                start = end;
            }
            return new ReadOnlyCollection<BaseHardCodedString>(instances);
        }

        /// <summary>
        /// Finds all instances of a text in the code files contained by the provided project
        /// </summary>
        /// <param name="project">Project to search</param>
        /// <param name="text">Text to look for</param>
        /// <returns>A collection of BaseHardCodedString implementations</returns>
        public static ReadOnlyCollection<BaseHardCodedString> FindAllInstancesInProject(Project project, string text)
        {
            Collection<BaseHardCodedString> instances = new Collection<BaseHardCodedString>();
            CodeFileCollection codeFiles = new CodeFileCollection(project);
            foreach (ProjectItem item in codeFiles)
            {
                foreach (BaseHardCodedString instance in FindAllInstancesInDocument(item, text))
                {
                    instances.Add(instance);
                }
            }
            return new ReadOnlyCollection<BaseHardCodedString>(instances);
        }

        /// <summary>
        /// Shortens a full namespace reference by looking at a list of namespaces that are imported in the code
        /// </summary>
        /// <param name="reference">Reference to shorten</param>
        /// <param name="namespaces">Collection of namespaces imported in the file</param>
        /// <returns>Shortest form the of the reference valid for the file</returns>
        /// <remarks>This implementation is same for C# and VB</remarks>
        public virtual string GetShortestReference(string reference, Collection<NamespaceImport> namespaces)
        {
            NamespaceImport longestMatch = null;
            // Retrieve the property part of the reference
            string property = reference;
            string[] parts = reference.Split('.');
            property = parts[parts.Length - 1];
            reference = reference.Substring(0, reference.Length - property.Length - (parts.Length == 0 ? 0 : 1));
            foreach (var ns in namespaces)
            {
                if (!reference.Equals(ns.NamespaceName) && reference.StartsWith(ns + ".") && ns.NamespaceName.Length > longestMatch.NamespaceName.Length)
                {
                    longestMatch = ns;
                }
            }
            if (longestMatch != null)
                reference = reference.Remove(0, longestMatch.NamespaceName.Length);
            if (reference.StartsWith("."))
            {
                reference = reference.Remove(0, 1);
            }
            if (longestMatch != null && longestMatch.Alias != null)
                reference = longestMatch.Alias + "." + longestMatch.NamespaceName;
            reference += "." + property;
            return reference;
        }

        /// <summary>
        /// Gets the TextDocument interface for the item provided
        /// </summary>
        /// <param name="item"></param>
        /// <returns></returns>
        public static TextDocument GetDocumentForItem(ProjectItem item)
        {
            if (item == null)
            {
                throw new ArgumentNullException("item");
            }
            if (item.Document == null)
            {
                item.Open(null);
            }
            return ((EnvDTE.TextDocument)item.Document.Object(null));
        }

        #endregion
    }

    /// <summary>
    /// Match result structure for hard coded string queries
    /// </summary>
    /// <remarks>FxCop rule is suppressed because this is a struct object</remarks>
    public struct MatchResult
    {
        private int startIndex;

        private int endIndex;

        private bool result;

        /// <summary>
        /// Starting index of the matched part (0 based indexing)
        /// </summary>
        public int StartIndex
        {
            get { return startIndex; }
            set { startIndex = value; }
        }

        /// <summary>
        /// Ending index of the matched part
        /// </summary>
        public int EndIndex
        {
            get { return endIndex; }
            set { endIndex = value; }
        }

        /// <summary>
        /// Result of the match, if false no match was found and StartIndex, EndIndex properties will be 0
        /// </summary>
        public bool Result
        {
            get { return result; }
            set { result = value; }
        }

        public override bool Equals(object obj)
        {
            if (obj is MatchResult)
            {
                MatchResult other = (MatchResult)obj;
                return this.Result == other.Result && this.StartIndex == other.StartIndex && this.EndIndex == other.EndIndex;
            }
            return base.Equals(obj);
        }

        public static bool operator ==(MatchResult obj1, MatchResult obj2)
        {
            return obj1.Equals(obj2);
        }

        public static bool operator !=(MatchResult obj1, MatchResult obj2)
        {
            return !obj1.Equals(obj2);
        }

        public override int GetHashCode()
        {
            return StartIndex + EndIndex + (result ? 0 : 1);
        }
    }
}
