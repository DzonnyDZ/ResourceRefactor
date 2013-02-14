/// Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using Microsoft.VSPowerToys.ResourceRefactor.Common;
using EnvDTE;
using System.IO;
using System.Reflection;

namespace Microsoft.VSPowerToys.ResourceRefactor.UnitTests
{
    /// <summary>
    /// Tests for checking if strings are recognized correctly in C# files
    /// </summary>
    [TestFixture]
    public class CSharpHardCodedStringRecognitionTests
    {

        private CSharpHardCodedString hardCodedString = new CSharpHardCodedString();

        #region String Recognition Tests

        /// <summary>
        /// These tests test CheckForHardCodedString with a simple line where string does not contain any special characters.
        /// Tests call the method with several different points of selection.
        /// </summary>
        #region Simple String Recognition Tests

        [Test]
        public void SimpleStringRecognitionInvalid1()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SimpleTestString, 1, 2);
            Assert.IsFalse(result.Result, Messages.MatchResultInvalid);
        }

        [Test]
        public void SimpleStringRecognitionInvalid2()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SimpleTestString, 15, 39);
            Assert.IsFalse(result.Result, Messages.MatchResultInvalid);
        }

        [Test]
        public void SimpleStringRecognitionInvalid3()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SimpleTestString, 15, 25);
            Assert.IsFalse(result.Result, Messages.MatchResultInvalid);
        }

        [Test]
        public void SimpleStringRecognitionInvalid4()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SimpleTestString, 19, 40);
            Assert.IsFalse(result.Result, Messages.MatchResultInvalid);
        }

        [Test]
        public void SimpleStringRecognitionValid1()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SimpleTestString, 20, 25);
            Assert.IsTrue(result.Result, Messages.MatchResultInvalid);
            Assert.AreEqual(19, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(39, result.EndIndex, Messages.MatchResultInvalid);
        }

        [Test]
        public void SimpleStringRecognitionValid2()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SimpleTestString, 19,25);
            Assert.IsTrue(result.Result, Messages.MatchResultInvalid);
            Assert.AreEqual(19, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(39, result.EndIndex, Messages.MatchResultInvalid);
        }

        [Test]
        public void SimpleStringRecognitionValid3()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SimpleTestString, 19,39);
            Assert.IsTrue(result.Result, Messages.MatchResultInvalid);
            Assert.AreEqual(19, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(39, result.EndIndex, Messages.MatchResultInvalid);
        }

        [Test]
        public void SimpleStringRecognitionValid4()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SimpleTestString, 25, 25);
            Assert.IsTrue(result.Result, Messages.MatchResultInvalid);
            Assert.AreEqual(19, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(39, result.EndIndex, Messages.MatchResultInvalid);
        }

#endregion

        #region Special Characters Recognition Tests

        [Test]
        public void SpecialCharacterRecognitionInvalid()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SpecialCharacterTestString, 3,5);
            Assert.IsFalse(result.Result, Messages.MatchResultInvalid);
        }

        [Test]
        public void SpecialCharacterRecognitionValid()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.SpecialCharacterTestString, 5, 5);
            Assert.IsTrue(result.Result, Messages.MatchResultInvalid);
            Assert.AreEqual(4, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(18, result.EndIndex, Messages.MatchResultInvalid);
        }
        #endregion

        /// <summary>
        /// This test uses a line which does not contain a valid hard coded string, no selection should return a match
        /// </summary>
        [Test]
        public void InvalidStringTest()
        {
            int length = TestStrings.InvalidTestString.Length;
            for (int i =0; i< length;i++) 
            {
                for (int j = i; j < length; j++) 
                {
                    MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.InvalidTestString, i, j);
                    Assert.IsFalse(result.Result, Messages.MatchResultInvalid);
                }
            }
        }

        #region Verbatim String Tests

        [Test]
        public void VerbatimStringTestValid1()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.VerbatimTestString, 3, 5);
            Assert.IsTrue(result.Result, Messages.MatchResultInvalid);
            Assert.AreEqual(2, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(20, result.EndIndex, Messages.MatchResultInvalid);
        }

        [Test]
        public void VerbatimStringTestValid2()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.VerbatimQuoteTestString, 3, 5);
            Assert.IsTrue(result.Result, Messages.MatchResultInvalid);
            Assert.AreEqual(2, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(9, result.EndIndex, Messages.MatchResultInvalid);
        }

        [Test]
        public void VerbatimInvalidStringTest()
        {
           MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.VerbatimInvalidTestString, 2,7);
           Assert.IsFalse(result.Result, Messages.MatchResultInvalid);
        }

        [Test]
        public void VerbatimMultiLineTestValid()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.VerbatimMultiLineTestString, 3, 10);
            Assert.IsTrue(result.Result, Messages.MatchResultInvalid);
            Assert.AreEqual(2, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(15, result.EndIndex, Messages.MatchResultInvalid);
        }

        [Test]
        public void VerbatimMultiLineTestInvalid()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.VerbatimMultiLineTestInvalidString, 3, 15);
            Assert.IsFalse(result.Result, Messages.MatchResultInvalid);
        }
        #endregion

        #endregion

    }

    /// <summary>
    /// Tests for checking functionality of CSharpHardCodedString object.
    /// </summary>
    [TestFixture]
    public class CSharpHardCodedStringFunctionalityTests
    {
        /// <summary>
        /// DTE object to be used to interface with Visual Studio
        /// </summary>
        private DTE extensibility;

        /// <summary>
        /// Project item for code file to use during tests.
        /// </summary>
        private ProjectItem codeFile;

        /// <summary>
        /// Resource file collection for the project being tested
        /// </summary>
        private ResourceFileCollection resources;

        /// <summary>
        /// Action object suitable for C# files
        /// </summary>
        private Common.IExtractResourceAction actionObject;

        /// <summary>
        /// Sets up the DTE object by creating an instance of Visual Studio
        /// </summary>
        [TestFixtureSetUp]
        public void FixtureSetup()
        {
            this.extensibility = SharedEnvironment.Instance;
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(3));
            this.codeFile = testProject.ProjectItems.Item("Program.cs");
            CommonMethods.CloseDocument(codeFile);
            actionObject = new Common.GenericCSharpExtractResourceAction();
            resources = new ResourceFileCollection(testProject, new FilterMethod(actionObject.IsValidResourceFile));
        }

        /// <summary>
        /// Tests Value property of CSharpHardCodedString object with different instances of strings.
        /// </summary>
        [Test]
        public void ValueGetterTest()
        {
            CSharpHardCodedString hcs = new CSharpHardCodedString(this.codeFile, 349, 362);
            Assert.AreEqual("Test String", hcs.Value, "Simple string Value get failed");
            hcs = new CSharpHardCodedString(this.codeFile, 383, 397);
            Assert.AreEqual("Test\"\"\n\t", hcs.Value, "String with escape characters Value get failed");
            hcs = new CSharpHardCodedString(this.codeFile, 418, 432);
            Assert.AreEqual(@"Test String", hcs.Value, "Value property for simple verbatim string failed");
            hcs = new CSharpHardCodedString(this.codeFile, 453,470);
            Assert.AreEqual(@"Test""String""", hcs.Value, "Value property for verbatim string with escaped quotes failed");
        }

        /// <summary>
        /// Tests Value property of CSharpHardCodedString object with different instances of strings.
        /// </summary>
        [Test]
        public void IndexGetterTests()
        {
            CSharpHardCodedString hcs = new CSharpHardCodedString(this.codeFile, 349, 362);
            TextDocument doc = ((EnvDTE.TextDocument)this.codeFile.Document.Object(null));
            EditPoint ep = doc.StartPoint.CreateEditPoint();
            string text = ep.GetLines(17, 18);
            Assert.AreEqual(text.IndexOf("\"Test\\\"\\\"\\n\\t\""), hcs.StartIndex);
            Assert.AreEqual(15, hcs.StartingLine);
        }

        /// <summary>
        /// Tests the "GetResourceReference" private method with several classes under different namespaces.
        /// </summary>
        [Test]
        public void GetResourceReferenceTest()
        {
            ResourceFile resFile = resources["Resource1.resx"];
            Assert.AreEqual("WindowsForms1.Resource1.Test", actionObject.GetResourceReference(resFile, "Test", null), "GetResourceReference does not work correctly under different namespaces");
        }

    }
}
