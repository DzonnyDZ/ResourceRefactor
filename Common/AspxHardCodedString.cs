using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using System.Web;
using EnvDTE;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common
{

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
        /// <param name="parent">Current object item (file)</param>
        /// <param name="start">Starting offset of the string</param>
        /// <param name="end">End offset of the string</param>
        /// <returns>A new instance of <see cref="AspxHardCodedString"/>.</returns>
        public override BaseHardCodedString CreateInstance(ProjectItem parent, int start, int end) {
            return new AspxHardCodedString(parent, start, end);
        }

        /// <summary>Gets the regular expression to identify strings in ASPX markup</summary>
        protected override string StringRegExp {
            get { return Strings.RegexAspxLiteral; }
        }

        /// <summary>Gets the regular expression object sting to identify ASPX comments</summary>
        protected override Regex CommentRegularExpression {
            get { return new Regex(Strings.RegexAspxComment); }
        }

        /// <summary>Cached value of the string</summary>
        /// <seealso cref="Value"/>
        private string value;

        /// <summary>Gets actual value of the string (without quotes and special characters)</summary>
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

        /// <summary>Returns a collection of namespaces imported in the files ('using' keyword in C#, or 'Imports' in VB.Net)</summary>
        /// <returns>This implementation returns an empty collection.</returns>
        /// <remarks>This list will be used to determine the replacement string</remarks>
        public override Collection<NamespaceImport> GetImportedNamespaces() {
            var namespaces = new Collection<NamespaceImport>();
            //GetWebConfigNamespaces(namespaces);
            //GetNamespacesFromFile(namespaces);
            return namespaces;
        }

        /// <summary>Shortens a full namespace reference by looking at a list of namespaces that are imported in the code</summary>
        /// <param name="reference">Reference to shorten</param>
        /// <param name="namespaces">Collection of namespaces imported in the file</param>
        /// <returns>Shortest form the of the reference valid for the file. This implementation just returns <paramref name="reference"/>.</returns>
        /// <remarks>This implementation does nothing, just returns <paramref name="reference"/>.</remarks>
        public override string GetShortestReference(string reference, Collection<NamespaceImport> namespaces) {
            return reference;
        }

    }
}
