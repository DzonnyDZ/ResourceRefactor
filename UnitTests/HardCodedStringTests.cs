/// Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Microsoft.VSPowerToys.ResourceRefactor.Common;
using EnvDTE;
using System.IO;
using System.Reflection;
using System.Collections.ObjectModel;

namespace Microsoft.VSPowerToys.ResourceRefactor.UnitTests
{
    /// <summary>
    /// Tests for methods in HardCodedString classes
    /// </summary>
    [TestFixture]
    public class BaseHardCodedStringTests
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
            extensibility.Solution.Close(false);
            extensibility.Solution.Open(Path.Combine(Environment.CurrentDirectory,
                Path.Combine(Paths.Default.ProjectFiles, "TestProject1\\TestProject1.sln")));
        }

        #region FindAllInstancesInDocument tests

        /// <summary>
        /// Tries to find a non existing string in a document
        /// </summary>
        [Test]
        public void FindInDocumentTestNoInstance()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(3));
            ProjectItem codeFile = testProject.ProjectItems.Item("Program.cs");
            ReadOnlyCollection<BaseHardCodedString> collection =
                BaseHardCodedString.FindAllInstancesInDocument(codeFile, "\"ggggg\"");
            Assert.AreEqual(0, collection.Count, Messages.CountInvalid);
        }

        /// <summary>
        /// Tests if "Test String" instances can be correctly found in the document.
        /// Also verifies location of instances are returned correctly
        /// </summary>
        [Test]
        public void FindInDocumentTest()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(3));
            ProjectItem codeFile = testProject.ProjectItems.Item("Program.cs");
            ReadOnlyCollection<BaseHardCodedString> collection =
                BaseHardCodedString.FindAllInstancesInDocument(codeFile, "\"Test String\"");
            Assert.AreEqual(3, collection.Count, Messages.CountInvalid);
            Assert.AreEqual(15, collection[0].StartingLine, Messages.MatchResultInvalid);
            Assert.AreEqual(17, collection[1].StartingLine, Messages.MatchResultInvalid);
            Assert.AreEqual(33, collection[2].StartingLine, Messages.MatchResultInvalid);
            Assert.AreEqual(19, collection[0].StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(32, collection[0].EndIndex, Messages.MatchResultInvalid);
        }

        /// <summary>
        /// Tests if comment detection works correctly for CSharp documents
        /// </summary>
        [Test]
        public void FindInDocumentCSharpCommentTest()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            ProjectItem codeFile = testProject.ProjectItems.Item("CommentTest.cs");
            ReadOnlyCollection<BaseHardCodedString> collection =
                BaseHardCodedString.FindAllInstancesInDocument(codeFile, "\"Test String\"");
            Assert.AreEqual(2, collection.Count, Messages.CountInvalid);
            Assert.AreEqual(11, collection[0].StartingLine, Messages.MatchResultInvalid);
            Assert.AreEqual(11, collection[1].StartingLine, Messages.MatchResultInvalid);
        }

        /// <summary>
        /// Tests if comment detection works correctly for VB documents
        /// </summary>
        [Test]
        public void FindInDocumentVBCommentTest()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(2));
            ProjectItem codeFile = testProject.ProjectItems.Item("Comment Test.vb");
            ReadOnlyCollection<BaseHardCodedString> collection =
                BaseHardCodedString.FindAllInstancesInDocument(codeFile, "\"TestString\"");
            Assert.AreEqual(1, collection.Count, Messages.CountInvalid);
        }
        #endregion

        #region FindAllInstancesInProject tests

        /// <summary>
        /// Tests if "Test String" instances can be correctly found in the project.
        /// There are 5 instances in the project, but 1 is in a designer file.
        /// </summary>
        [Test]
        public void FindInCSharpProject()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(3));
            ReadOnlyCollection<BaseHardCodedString> collection =
                BaseHardCodedString.FindAllInstancesInProject(testProject, "\"Test String\"");
            Assert.AreEqual(4, collection.Count, Messages.CountInvalid);
        }

        /// <summary>
        /// Tests if "Test String" instances can be correctly found in the project.
        /// There are 2 instances in the project, but 1 is in a C# file and should be ignored
        /// </summary>
        [Test]
        public void FindInVBProject()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(2));
            ReadOnlyCollection<BaseHardCodedString> collection =
                BaseHardCodedString.FindAllInstancesInProject(testProject, "\"Test String\"");
            Assert.AreEqual(1, collection.Count, Messages.CountInvalid);
        }

        /// <summary>
        /// Tests if "Test String" instances can be correctly found in the project.
        /// There are 2 instances in the project in both a C# and a VB file
        /// </summary>
        [Test]
        public void FindInWebProject()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(4));
            ReadOnlyCollection<BaseHardCodedString> collection =
                BaseHardCodedString.FindAllInstancesInProject(testProject, "\"Test String\"");
            Assert.AreEqual(2, collection.Count, Messages.CountInvalid);
        }

        #endregion
    }

    /// <summary>
    /// Contains tests for GetShortestReference method
    /// </summary>
    [TestFixture]
    public class GetShortestReferenceTests
    {
        #region GetShortestReference method tests
        /// <summary>
        /// Tests GetShortestReference method with an empty list of references
        /// </summary>
        [Test]
        public void GetShortestReferenceTestEmptyList()
        {
            Collection<string> namespaces = new Collection<string>();
            string result = GetShortestReferenceWrapper("System.Windows.Forms", namespaces);
            Assert.AreEqual("System.Windows.Forms", result);
        }

        /// <summary>
        /// Tests GetShortestReference method when no namespace shares the same prefix as the reference
        /// </summary>
        [Test]
        public void GetShortestReferenceTestNoPrefix()
        {
            Collection<string> namespaces = new Collection<string>();
            namespaces.Add("Microsoft.VS");
            namespaces.Add("Test1.Test2");
            namespaces.Add("Test2");
            string result = GetShortestReferenceWrapper("System.Windows.Forms.TextBox", namespaces);
            Assert.AreEqual("System.Windows.Forms.TextBox", result);
        }

        /// <summary>
        /// Tests GetShortestReference method when there is a shorter reference possible
        /// </summary>
        [Test]
        public void GetShortestReferenceTestPrefix()
        {
            Collection<string> namespaces = new Collection<string>();
            namespaces.Add("Microsoft.VS");
            namespaces.Add("System.Windows");
            namespaces.Add("System");
            namespaces.Add("Test2");
            string result = GetShortestReferenceWrapper("System.Windows.Forms.TextBox", namespaces);
            Assert.AreEqual("Forms.TextBox", result);
        }

        /// <summary>
        /// Tests the scenario where reference is included in the namespaces list as well. Correct functionality
        /// should ignore that string
        /// </summary>
        [Test]
        public void GetShortestReferenceTestFullTextIncluded()
        {
            Collection<string> namespaces = new Collection<string>();
            namespaces.Add("Microsoft.VS");
            namespaces.Add("System.Windows.Forms.TextBox");
            namespaces.Add("System");
            string result = GetShortestReferenceWrapper("System.Windows.Forms.TextBox", namespaces);
            Assert.AreEqual("Windows.Forms.TextBox", result);
        }

        /// <summary>
        /// A wrapper for the protected method: BaseHardCodedString.GetShortestReference
        /// </summary>
        /// <param name="reference"></param>
        /// <param name="namespaces"></param>
        /// <returns></returns>
        private static string GetShortestReferenceWrapper(string reference, Collection<string> namespaces)
        {
            Type baseType = ((Common.BaseHardCodedString)new Common.CSharpHardCodedString()).GetType();
            MethodInfo m = baseType.GetMethod("GetShortestReference",
                BindingFlags.Static | BindingFlags.Public | BindingFlags.FlattenHierarchy | BindingFlags.InvokeMethod);
            return (String)(m.Invoke(null, new Object[] { reference, namespaces }));
        }

        #endregion
    }
}
