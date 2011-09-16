using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace Microsoft.VSPowerToys.ResourceRefactor.Common {
    /// <summary>Represents imported namespace</summary>
    [DebuggerDisplay("{OriginalString}")]
    public class NamespaceImport {
        /// <summary>CTor - creates a new instance of the <see cref="NamespaceImport"/> class</summary>
        /// <param name="originalString">Origibal text of namespace import statement</param>
        /// <param name="namespaceName">Name of the namespace</param>
        /// <param name="alias">Namespace alias (if specified)</param>
        /// <exception cref="ArgumentNullException"><paramref name="originalString"/> or <paramref name="namespaceName"/> is null</exception>
        public NamespaceImport(string originalString, string namespaceName, string alias = null) {
            if (originalString == null) throw new ArgumentNullException("originalString");
            if (namespaceName == null) throw new ArgumentNullException("namespaceName");
            this.originalString = originalString;
            this.namespaceName = namespaceName;
            this.alias = alias;
        }
        private string originalString;
        private string namespaceName;
        private string alias;
        /// <summary>Returns a <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.</summary>
        /// <returns>A <see cref="T:System.String"/> that represents the current <see cref="T:System.Object"/>.</returns>
        /// <filterpriority>2</filterpriority>
        public override string ToString() { return OriginalString; }
        /// <summary>Gets original representation of namespace import statement as specified in code</summary>
        public string OriginalString { get { return originalString; } }
        /// <summary>Gets name of namespace imported</summary>
        public string NamespaceName { get { return namespaceName; } }
        /// <summary>Gets alias assigned to imported namespace</summary>
        public string Alias { get { return alias; } }
    }
}
