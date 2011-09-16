/// Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using EnvDTE;
using System.IO;
using Microsoft.VSPowerToys.ResourceRefactor.Common;

namespace Microsoft.VSPowerToys.ResourceRefactor.UnitTests
{
    [TestFixture]
    public class CodeFileCollectionTests
    {

        /// <summary>
        /// DTE object to be used to interface with Visual Studio
        /// </summary>
        private DTE extensibility;

        /// <summary>
        /// Sets up the DTE object by creating an instance of Visual Studio
        /// </summary>
        [TestFixtureSetUp]
        public void FixtureSetup()
        {
            this.extensibility = SharedEnvironment.Instance;
        }
        
        #region CodeFileCollection Tests

        /// <summary>
        /// Tests if ResourceFileCollection correctly gathers resource file information from a simple C# project
        /// </summary>
        [Test]
        public void CodeFileCollectionSimpleProjectTest()
        {
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            CodeFileCollection collection = new CodeFileCollection(testProject);
            String[] names = new string[] { "AssemblyInfo.cs", "CommentTest.cs"};
            this.MatchArrayAndCollection(names, collection);
        }

        /// <summary>
        /// Tests if ResourceFileCollection correctly gathers resource file information from a VB project with UI elements
        /// </summary>
        [Test]
        public void CodeFileCollectionVBProject()
        {
            Project testProject = (Project)(extensibility.Solution.Projects.Item(2));
            CodeFileCollection collection = new CodeFileCollection(testProject);
            String[] names = new string[] { "Form1.vb", "AssemblyInfo.vb", "Comment Test.vb"};
            this.MatchArrayAndCollection(names, collection);
        }

        /// <summary>
        /// Tests if ResourceFileCollection correctly gathers resource file information from a C# project with UI elements and a resource file
        /// in a directory
        /// </summary>
        [Test]
        public void CodeFileCollectionCSharpProject()
        {
            Project testProject = (Project)(extensibility.Solution.Projects.Item(3));
            CodeFileCollection collection = new CodeFileCollection(testProject);
            String[] names = new string[] { "AssemblyInfo.cs", "Form1.cs", "Program.cs" };
            this.MatchArrayAndCollection(names, collection);
        }

        /// <summary>
        /// Tests if ResourceFileCollection correctly gathers resource file information from a web project
        /// </summary>
        /// <remarks>Web project tested contains an invalid resource file outside of App_GlobalResources directory</remarks>
        [Test]
        public void CodeFileCollectionWebProject()
        {
            Project testProject = (Project)(extensibility.Solution.Projects.Item(4));
            CodeFileCollection collection = new CodeFileCollection(testProject);
            String[] names = new string[] {"Class1.vb", "Default.aspx.cs" };
            this.MatchArrayAndCollection(names, collection);
        }

        /// <summary>
        /// Tests if array of file names matches to the files in the collection
        /// </summary>
        /// <param name="names"></param>
        /// <param name="collection"></param>
        private void MatchArrayAndCollection(String[] names, CodeFileCollection collection)
        {
            Assert.AreEqual(names.Length, collection.Count, Messages.CodeFilesCountInvalid);
            foreach (string name in names)
            {
                Assert.IsNotNull(collection.GetCodeFile(name), Messages.CodeFileNotFound);
            }
        }
        #endregion
    }
}
