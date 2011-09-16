/// Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using EnvDTE;
using System.IO;
using Microsoft.VSPowerToys.ResourceRefactor.Common;
using System.Collections.ObjectModel;

namespace Microsoft.VSPowerToys.ResourceRefactor.UnitTests
{
    [TestFixture]
    public class ResourceFileTests
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
            extensibility = SharedEnvironment.Instance;
        }

        #region ResourceFileCollection Tests

        /// <summary>
        /// Tests if ResourceFileCollection correctly gathers resource file information from a simple C# project
        /// </summary>
        [Test]
        public void ResourceFileCollectionSimpleProjectTest()
        {
            // Get Project object
            Common.IExtractResourceAction actionObject = new Common.GenericCSharpExtractResourceAction();
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(actionObject.IsValidResourceFile));
            Assert.AreEqual(2, collection.Count, Messages.ResourceFilesCountInvalid);
            Assert.AreEqual(collection[0].DisplayName, "Resource1.resx", Messages.ResourceFileNotFound);
        }

        /// <summary>
        /// Tests if ResourceFileCollection correctly gathers resource file information from a VB project with UI elements
        /// </summary>
        [Test]
        public void ResourceFileCollectionVBProject()
        {
            // Get Project object
            Common.IExtractResourceAction actionObject = new Common.GenericVBExtractResourceAction();
            Project testProject = (Project)(extensibility.Solution.Projects.Item(2));
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(actionObject.IsValidResourceFile));
            String[] names = new string[] { "Resource1.resx", "(Default resources)" };
            Assert.AreEqual(names.Length, collection.Count, Messages.ResourceFilesCountInvalid);
            foreach (string name in names)
            {
                Assert.IsNotNull(collection.GetResourceFile(name), Messages.ResourceFileNotFound);
            }
        }

        /// <summary>
        /// Tests if ResourceFileCollection correctly gathers resource file information from a C# project with UI elements and a resource file
        /// in a directory
        /// </summary>
        [Test]
        public void ResourceFileCollectionCSharpProject()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(3));
            IExtractResourceAction action = new GenericCSharpExtractResourceAction();
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(action.IsValidResourceFile));
            String[] names = new string[] { "Resource1.resx", "Test1.resx", "Resources.resx" };
            Assert.AreEqual(names.Length, collection.Count, Messages.ResourceFilesCountInvalid);
            foreach (string name in names)
            {
                Assert.IsNotNull(collection.GetResourceFile(name), Messages.ResourceFileNotFound);
            }
        }

        /// <summary>
        /// Tests if ResourceFileCollection correctly gathers resource file information from a web project
        /// </summary>
        /// <remarks>Web project tested contains an invalid resource file outside of App_GlobalResources directory</remarks>
        [Test]
        public void ResourceFileCollectionWebProject()
        {
            Common.IExtractResourceAction actionObject = new Common.WebsiteCSharpExtractResourceAction();
            Project testProject = (Project)(extensibility.Solution.Projects.Item(4));
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(actionObject.IsValidResourceFile));
            String[] names = new string[] { "Resource.resx" };
            Assert.AreEqual(names.Length, collection.Count, Messages.ResourceFilesCountInvalid);
            foreach (string name in names)
            {
                Assert.IsNotNull(collection.GetResourceFile(name), Messages.ResourceFileNotFound);
            }
        }

        /// <summary>
        /// Tests if resource file collection works correctly if a project item's file
        /// is missing.
        /// </summary>
        [Test]
        public void ResourceFileCollectionMissingFile()
        {
            Common.IExtractResourceAction actionObject = new Common.GenericCSharpExtractResourceAction();
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            string filePath = Path.Combine(Paths.Default.ProjectFiles, @"TestProject1\TestProject1\Resource1.resx");
            bool readOnly = CommonMethods.ToggleReadOnly(filePath, true);
            File.Move(filePath, filePath + ".bak");
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(actionObject.IsValidResourceFile));
            Assert.AreEqual(1, collection.Count, Messages.ResourceFilesCountInvalid);
            File.Move(filePath + ".bak", filePath);
            if (readOnly)
            {
                CommonMethods.ToggleReadOnly(filePath, false);
            }
        }

        #endregion

        #region ResourceFile Related Tests

        /// <summary>
        /// Tests if resource file contents are read correctly initially.
        /// </summary>
        [Test]
        public void ResourceFileInitialReadTest()
        {
            Common.IExtractResourceAction actionObject = new Common.GenericCSharpExtractResourceAction();
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(actionObject.IsValidResourceFile));
            ResourceFile testFile = collection[0];
            Assert.AreEqual("ResXFileCodeGenerator", testFile.CustomToolName, "Custom tool name is invalid");
            Assert.AreEqual(2, testFile.Resources.Count, Messages.ResourceCountInvalid);
            Assert.AreEqual("Test", testFile.GetValue("TestResource"), Messages.ResourceValueInvalid);
        }

        /// <summary>
        /// Tests if resources added by AddResource method are ignored correctly when Refresh method is called
        /// before save.
        /// </summary>
        [Test]
        public void ResourceFileReadAfterRefreshWithoutSave()
        {
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            IExtractResourceAction action = new GenericCSharpExtractResourceAction();
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(action.IsValidResourceFile));
            ResourceFile testFile = collection[0];
            Assert.AreEqual(2, testFile.Resources.Count, Messages.ResourceCountInvalid);
            Assert.AreEqual("Test", testFile.GetValue("TestResource"), Messages.ResourceValueInvalid);
            testFile.AddResource("Test2", "Test3", "Comment Test");
            testFile.Refresh();
            Assert.AreEqual(2, testFile.Resources.Count, Messages.ResourceCountInvalid);
            Assert.AreEqual("Test", testFile.GetValue("TestResource"), Messages.ResourceValueInvalid);
        }

        /// <summary>
        /// Tests if resource file contents are saved correct after AddResource method is called.
        /// </summary>
        [Test]
        public void ResourceFileReadAfterSave()
        {
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            IExtractResourceAction action = new GenericCSharpExtractResourceAction();
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(action.IsValidResourceFile));
            ResourceFile testFile = collection[1];
            FileAttributes attributes = File.GetAttributes(testFile.FileName);
            if (File.Exists(testFile.FileName + ".bak"))
            {
                File.SetAttributes(testFile.FileName + ".bak", attributes & ~FileAttributes.ReadOnly);
                File.Delete(testFile.FileName + ".bak");
            }
            File.Copy(testFile.FileName, testFile.FileName + ".bak");
            File.SetAttributes(testFile.FileName, attributes & ~FileAttributes.ReadOnly);
            File.SetAttributes(testFile.FileName + ".bak", attributes & ~FileAttributes.ReadOnly);
            try
            {
                Assert.AreEqual(1, testFile.Resources.Count, Messages.ResourceCountInvalid);
                Assert.AreEqual("Test", testFile.GetValue("TestString"), Messages.ResourceValueInvalid);
                testFile.AddResource("Test2", "Test3", "Comment Test");
                testFile.SaveFile();
                testFile.Refresh();
                Assert.AreEqual(2, testFile.Resources.Count, Messages.ResourceCountInvalid);
                Assert.AreEqual("Test", testFile.GetValue("TestString"), Messages.ResourceValueInvalid);
                Assert.AreEqual("Comment Test", testFile.Resources["TestString"].Comment, Messages.ResourceCommentInvalid);
                Assert.AreEqual("Test3", testFile.GetValue("Test2"), Messages.ResourceValueInvalid);
                Assert.AreEqual("Comment Test", testFile.Resources["Test2"].Comment, Messages.ResourceCommentInvalid);
            }
            finally
            {
                File.Delete(testFile.FileName);
                File.Move(testFile.FileName + ".bak", testFile.FileName);
                File.SetAttributes(testFile.FileName, attributes);
            
            }
        }

        /// <summary>
        /// Tests if ResourceFile class correctly throws an exception when adding multiple resources
        /// with the same name
        /// </summary>
        [Test]
        [ExpectedException("System.ArgumentException")]
        public void ResourceFileMultipleAddsSameName()
        {
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            IExtractResourceAction action = new GenericCSharpExtractResourceAction();
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(action.IsValidResourceFile));
            ResourceFile testFile = collection[0];
            testFile.AddResource("Test2", "Test3", "Comment Test");
            testFile.AddResource("Test2", "TestTest", "Comment");
        }

        /// <summary>
        /// Tests if ResourceFile class correctly throws an exception when adding multiple resources
        /// with the same name but with case difference
        /// </summary>
        [Test]
        [ExpectedException("System.ArgumentException")]
        public void ResourceFileMultipleAddsCaseDifference()
        {
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            IExtractResourceAction action = new GenericCSharpExtractResourceAction();
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(action.IsValidResourceFile));
            ResourceFile testFile = collection[0];
            testFile.AddResource("Test2", "Test3", "Comment Test");
            testFile.AddResource("test2", "TestTest", "Comment");
        }

        /// <summary>
        /// Tests GetAllMatches method in the resource file object.
        /// One of the strings should match exactly to the string searched in the resource file.
        /// </summary>
        [Test]
        public void GetAllMatchesTest()
        {
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            Common.IExtractResourceAction action = new GenericCSharpExtractResourceAction();
            ResourceFileCollection collection = new ResourceFileCollection(testProject, new FilterMethod(action.IsValidResourceFile));
            ResourceFile testFile = collection["Resource1.resx"];
            Collection<ResourceMatch> matches = testFile.GetAllMatches("Test");
            foreach (ResourceMatch m in matches)
            {
                if (m.Percentage == StringMatch.ExactMatch)
                {
                    Assert.AreEqual("TestResource", m.ResourceName, "Did not match to expected resource entry");
                    Assert.AreEqual("Test", m.Value, "Value does not match the string searched");
                }
            }
        }
        #endregion

    }
}
