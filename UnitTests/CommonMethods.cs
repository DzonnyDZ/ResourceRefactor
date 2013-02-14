using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NUnit.Framework;
namespace Microsoft.VSPowerToys.ResourceRefactor.UnitTests
{
    class CommonMethods
    {
        /// <summary>
        /// Toggles read only property of a file
        /// </summary>
        /// <param name="fileName">Full path to the file name</param>
        /// <param name="readOnly">value of read only property</param>
        /// <returns>true if status was changed as a result of the operation</returns>
        public static bool ToggleReadOnly(string fileName, bool readOnly)
        {
            FileAttributes attributes = File.GetAttributes(fileName);
            bool returnValue = false;
            if (readOnly)
            {
                returnValue = !((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly);
                attributes = attributes | FileAttributes.ReadOnly;
            }
            else
            {
                returnValue = ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly);
                attributes &= ~FileAttributes.ReadOnly;
            }
            File.SetAttributes(fileName, attributes);
            return returnValue;
        }

        /// <summary>
        /// Closes the document to undo changes if it is open
        /// </summary>
        /// <param name="codeFile">Document to close</param>
        internal static void CloseDocument(EnvDTE.ProjectItem codeFile)
        {
            if (codeFile.Document != null)
            {
                codeFile.Document.Close(EnvDTE.vsSaveChanges.vsSaveChangesNo);
            }
        }
    }

    /// <summary>
    /// A wrapper for sharing DTE environment between tests
    /// </summary>
    internal sealed class SharedEnvironment : IDisposable
    {
        private EnvDTE.DTE extensibility;

        private static SharedEnvironment instance;

        /// <summary>
        /// Gets the shared Visual Studio instance
        /// </summary>
        public static EnvDTE.DTE Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new SharedEnvironment();
                }
                return instance.extensibility;
            }
        }

        /// <summary>
        /// Gets the DTE object for the current instance
        /// </summary>
        public EnvDTE.DTE Extensibility
        {
            get { return this.extensibility; }
        }

        /// <summary>
        /// Creates a new DTE environment
        /// </summary>
        public SharedEnvironment()
        {
            Type dteType = Type.GetTypeFromProgID("VisualStudio.DTE.10.0");
            Assert.IsNotNull(dteType, "Visual Studio DTE Type could not be found");
            extensibility = (EnvDTE.DTE)(System.Activator.CreateInstance(dteType));
            extensibility.MainWindow.Visible = false;
            MessageFilter.Register();
            extensibility.Solution.Open(Path.Combine(Environment.CurrentDirectory,
                Path.Combine(Paths.Default.ProjectFiles, "TestProject1\\TestProject1.sln")));
        }

        /// <summary>
        /// Dispose of non managed resources
        /// </summary>
        public void Dispose()
        {
            Dispose(false);
        }

        /// <summary>
        /// Dispose of resources
        /// </summary>
        /// <param name="disposing">If true managed resources will be disposed as well</param>
        private void Dispose(bool disposing)
        {
            MessageFilter.Revoke();
            if (extensibility != null)
            {
                extensibility.Quit();
            }
        }
    }
}
