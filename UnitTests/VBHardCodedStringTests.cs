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
    public class VBHardCodedStringRecognitionTests
    {

        private VBHardCodedString hardCodedString = new VBHardCodedString();

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

        /// <summary>
        /// This test uses a line which does not contain a valid hard coded string, no selection should return a match
        /// </summary>
        [Test]
        public void InvalidStringTest()
        {
            int length = TestStrings.InvalidTestVBString.Length;
            for (int i =0; i< length;i++) 
            {
                for (int j = i; j < length; j++) 
                {
                    MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.InvalidTestVBString, i, j);
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
            Assert.AreEqual(3, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(20, result.EndIndex, Messages.MatchResultInvalid);
        }

        [Test]
        public void VerbatimStringTestValid2()
        {
            MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.VerbatimQuoteTestString, 3, 5);
            Assert.IsTrue(result.Result, Messages.MatchResultInvalid);
            Assert.AreEqual(3, result.StartIndex, Messages.MatchResultInvalid);
            Assert.AreEqual(9, result.EndIndex, Messages.MatchResultInvalid);
        }

        [Test]
        public void VerbatimInvalidStringTest()
        {
           MatchResult result = hardCodedString.CheckForHardCodedString(TestStrings.VerbatimInvalidTestString, 2,7);
           Assert.IsFalse(result.Result, Messages.MatchResultInvalid);
        }
        #endregion

        #endregion

    }

    /// <summary>
    /// Tests for checking functionality of VBHardCodedString object.
    /// </summary>
    [TestFixture]
    public class VBHardCodedStringFunctionalityTests
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
        /// Sets up the DTE object by creating an instance of Visual Studio
        /// </summary>
        [TestFixtureSetUp]
        public void FixtureSetup()
        {
            this.extensibility = SharedEnvironment.Instance;
            // Get Project object
            Project testProject = (Project)(extensibility.Solution.Projects.Item(2));
            IExtractResourceAction action = new GenericVBExtractResourceAction();
            this.codeFile = testProject.ProjectItems.Item("Form1.vb");
            CommonMethods.CloseDocument(codeFile);
            resources = new ResourceFileCollection(testProject, new FilterMethod(action.IsValidResourceFile));
        }

        /// <summary>
        /// Tests Value property of VBHardCodedString object with different instances of strings.
        /// </summary>
        [Test]
        public void ValueGetterTest()
        {
            VBHardCodedString hcs = new VBHardCodedString(this.codeFile, 157, 172);
            Assert.AreEqual("Test Instance", hcs.Value, "Simple string Value get failed");
            Assert.AreEqual("\"Test Instance\"", hcs.RawValue, "Raw Value get failed");
            hcs = new VBHardCodedString(this.codeFile,  188,204);
            Assert.AreEqual(@"Test""Instance", hcs.Value, "Value property for verbatim string with escaped quotes failed");
            Assert.AreEqual("\"Test\"\"Instance\"", hcs.RawValue, "Raw Value get failed");
            
        }

        /// <summary>
        /// Tests Value property of VBHardCodedString object with different instances of strings.
        /// </summary>
        [Test]
        public void IndexGetterTests()
        {
            VBHardCodedString hcs = new VBHardCodedString(this.codeFile, 157, 172);
            TextDocument doc = ((EnvDTE.TextDocument)this.codeFile.Document.Object(null));
            EditPoint ep = doc.StartPoint.CreateEditPoint();
            string text = ep.GetLines(9,10);
            Assert.AreEqual(text.IndexOf("\"Test Instance\""), hcs.StartIndex);
            Assert.AreEqual(text.IndexOf("\"Test Instance\"") + "\"Test Instance\"".Length, hcs.EndIndex);
            Assert.AreEqual(8, hcs.StartingLine);
        }

        /// <summary>
        /// Tests the "GetResourceReference" private method with several classes
        /// </summary>
        [Test]
        public void GetResourceReferenceTest()
        {
            ResourceFile resFile = resources["Resource1.resx"];
            VBHardCodedString hcs = new VBHardCodedString(this.codeFile, 157, 172);
            IExtractResourceAction action = new GenericVBExtractResourceAction();
            Assert.AreEqual("My.Resources.Resource1.Test", action.GetResourceReference(resFile, "Test"), "GetResourceReference does not work correctly in VB");
        }

        
    }
}
