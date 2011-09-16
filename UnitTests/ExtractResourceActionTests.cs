using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using EnvDTE;
using Microsoft.VSPowerToys.ResourceRefactor.Common;
using System.IO;

namespace Microsoft.VSPowerToys.ResourceRefactor.UnitTests
{
    /// <summary>
    /// Contains tests for IExtractResourceAction implementations
    /// </summary>
    [TestFixture]
    public class ExtractResourceActionTests
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

        #region Generic C# project

        /// <summary>
        /// Tests the Replace method, since replace method does not save the results this method creates another edit point before the string to read the whole line 
        /// to check if it matches to expected output.
        /// </summary>
        [Test]
        public void GenericCSharpReplaceMethodTest()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(3));
            ProjectItem codeFile = testProject.ProjectItems.Item("Program.cs");
            CSharpHardCodedString hcs = new CSharpHardCodedString(codeFile, 349, 362);
            TestReplaceMethod(codeFile, "Resource1.resx", hcs, TestStrings.CSharpReplaceTestExpectedLine, "Test");
         }

        /// <summary>
        /// Tests the Replace method when the file is read only
        /// </summary>
        [Test]
        [ExpectedException(typeof(Common.FileReadOnlyException))]
        public void GenericCSharpReplaceMethodReadOnlyTest()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(3));
            ProjectItem codeFile = testProject.ProjectItems.Item("Program.cs");
            CSharpHardCodedString hcs = new CSharpHardCodedString(codeFile, 19, 32);
            IExtractResourceAction actionObject = new Common.GenericCSharpExtractResourceAction();
            ResourceFileCollection resources = new ResourceFileCollection(testProject, new FilterMethod(actionObject.IsValidResourceFile));

            ResourceFile resFile = resources["Resource1.resx"];
            string fileName = codeFile.get_FileNames(1);
            FileAttributes oldAttributes = System.IO.File.GetAttributes(fileName);
            System.IO.File.SetAttributes(fileName, oldAttributes | FileAttributes.ReadOnly);
            try
            {
                ExtractToResourceActionSite refactorSite = new ExtractToResourceActionSite(hcs);
                refactorSite.ExtractStringToResource(resFile, "Test");
            }
            finally
            {
                System.IO.File.SetAttributes(fileName, oldAttributes);
            }
        }

        #endregion

        #region Generic VB project

        /// <summary>
        /// Tests the Replace method, since replace method does not save the results this method creates another edit point before the string to read the whole line 
        /// to check if it matches to expected output.
        /// </summary>
        [Test]
        public void GenericVBReplaceMethodTest()
        {

            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(2));
            ProjectItem codeFile = testProject.ProjectItems.Item("Form1.vb");
            VBHardCodedString hcs = new VBHardCodedString(codeFile, 157, 172);
            TestReplaceMethod(codeFile, "Resource1.resx", hcs, TestStrings.VBReplaceTestExpectedLine, "Test");
        }

        #endregion

        /// <summary>
        /// Tests the Replace method, since replace method does not save the results this method creates another edit point before the string to read the whole line 
        /// to check if it matches to expected output.
        /// </summary>
        [Test]
        public void CSharpWebsiteReplaceMethodTest()
        {

            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(4));
            ProjectItem codeFile = testProject.ProjectItems.Item("Default.aspx.cs");
            BaseHardCodedString hcs = new CSharpHardCodedString(codeFile, 376, 389);
            TestReplaceMethod(codeFile, "Resource.resx", hcs, TestStrings.CSharpWebsiteTestString, "TestResource");                    
        }

        /// <summary>
        /// Tests the Replace method, since replace method does not save the results this method creates another edit point before the string to read the whole line 
        /// to check if it matches to expected output.
        /// </summary>
        [Test]
        public void VBWebsiteReplaceMethodTest()
        {

            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(4));
            ProjectItem codeFile = testProject.ProjectItems.Item("App_Code").ProjectItems.Item("Class1.vb");
            BaseHardCodedString hcs = new CSharpHardCodedString(codeFile, 73, 86);
            TestReplaceMethod(codeFile, "Resource.resx", hcs, TestStrings.VBWebsiteTestString, "TestResource");
        }

        /// <summary>
        /// Invokes replace method on the provided objects
        /// </summary>
        /// <param name="codeFile">Code file to test</param>
        /// <param name="resFile">Resource file to use</param>
        /// <param name="hcs">Hardcoded string to use</param>
        /// <param name="expected">Expected test string</param>
        internal static void TestReplaceMethod(ProjectItem codeFile, string resourceFileName, BaseHardCodedString hcs, string expected, string resourceName)
        {
            string fileName = codeFile.get_FileNames(1);
            bool readOnly = false;
            try
            {
                readOnly = CommonMethods.ToggleReadOnly(fileName, false);
                ExtractToResourceActionSite refactorSite = new ExtractToResourceActionSite(hcs);
                ResourceFileCollection resources = new ResourceFileCollection(codeFile.ContainingProject,
                    new FilterMethod(refactorSite.ActionObject.IsValidResourceFile));
                ResourceFile resFile = resources[resourceFileName];
                refactorSite.ExtractStringToResource(resFile, resourceName);
                TextDocument doc = ((EnvDTE.TextDocument)codeFile.Document.Object(null));
                EditPoint ep = doc.StartPoint.CreateEditPoint();
                string line = ep.GetLines(hcs.StartingLine + 1, hcs.StartingLine + 2);
                Assert.AreEqual(expected, line, "New line does not match the expected output");
            }
            finally
            {
                if (readOnly) CommonMethods.ToggleReadOnly(fileName, true);
            }
        }
    }

    /// <summary>
    /// Contains tests for IExtractResourceAction implementations designed to work
    /// with WebApplications
    /// </summary>
    [TestFixture]
    public class WebApplicationTests
    {
        /// <summary>
        /// DTE object to be used to interface with Visual Studio
        /// </summary>
        private DTE extensibility;

        private SharedEnvironment environment;

        /// <summary>
        /// Sets up the DTE object by creating an instance of Visual Studio
        /// </summary>
        [TestFixtureSetUp]
        public void FixtureSetup()
        {
            this.environment = new SharedEnvironment();
            this.extensibility = this.environment.Extensibility;
            extensibility.Solution.Open(Path.Combine(Environment.CurrentDirectory,
                Path.Combine(Paths.Default.ProjectFiles, "WebApplication\\WebApplication.sln")));
            if (extensibility.Solution.Projects.Item(1).ProjectItems == null)
            {
                Assert.Ignore("Web application add-ins are not installed.");
            }
        }

        /// <summary>
        /// Closes the instance of Visual Studio created during the test
        /// </summary>
        [TestFixtureTearDown]
        public void FixtureTearDown()
        {
            if (this.environment != null)
            {
                this.environment.Dispose();
            }
        }

        /// <summary>
        /// Tests the Replace method, since replace method does not save the results this method creates another edit point before the string to read the whole line 
        /// to check if it matches to expected output.
        /// </summary>
        [Test]
        public void CSharpReplaceTestGlobalResource()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            ProjectItem codeFile = testProject.ProjectItems.Item("Test.cs");
            CommonMethods.CloseDocument(codeFile);
            CSharpHardCodedString hcs = new CSharpHardCodedString(codeFile, 307, 313);
            ExtractResourceActionTests.TestReplaceMethod(codeFile, "Global.resx", hcs, TestStrings.WebApplicationGlobalTestStringCSharp, "Test");
        }

        /// <summary>
        /// Tests the Replace method, since replace method does not save the results this method creates another edit point before the string to read the whole line 
        /// to check if it matches to expected output.
        /// </summary>
        [Test]
        public void CSharpReplaceTestLocalResource()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(1));
            ProjectItem codeFile = testProject.ProjectItems.Item("Test.cs");
            CommonMethods.CloseDocument(codeFile);
            CSharpHardCodedString hcs = new CSharpHardCodedString(codeFile, 307, 313);
            ExtractResourceActionTests.TestReplaceMethod(codeFile, "LocalResource.resx", hcs, TestStrings.WebApplicationLocalTestStringCSharp, "Test");
        }

        /// <summary>
        /// Tests the Replace method, since replace method does not save the results this method creates another edit point before the string to read the whole line 
        /// to check if it matches to expected output.
        /// </summary>
        [Test]
        public void VBReplaceTestGlobalResource()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(2));
            ProjectItem codeFile = testProject.ProjectItems.Item("Test.vb");
            CommonMethods.CloseDocument(codeFile);
            CSharpHardCodedString hcs = new CSharpHardCodedString(codeFile, 40, 46);
            ExtractResourceActionTests.TestReplaceMethod(codeFile, "Global.resx", hcs, TestStrings.WebApplicationGlobalTestStringVB, "Test");
        }

        /// <summary>
        /// Tests the Replace method, since replace method does not save the results this method creates another edit point before the string to read the whole line 
        /// to check if it matches to expected output.
        /// </summary>
        [Test]
        public void VBReplaceTestLocalResource()
        {
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(2));
            ProjectItem codeFile = testProject.ProjectItems.Item("Test.vb");
            CommonMethods.CloseDocument(codeFile);
            CSharpHardCodedString hcs = new CSharpHardCodedString(codeFile, 40, 46);
            ExtractResourceActionTests.TestReplaceMethod(codeFile, "LocalResource.resx", hcs, TestStrings.WebApplicationLocalTestStringVB, "Test");
        }
    }
}
