/// Copyright (c) Microsoft Corporation.  All rights reserved.
using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.VSPowerToys.ResourceRefactor.Common;
using NUnit.Framework;

namespace Microsoft.VSPowerToys.ResourceRefactor.UnitTests
{
    /// <summary>
    /// Tests methods and operations on Common.MatchResult object
    /// </summary>
    [TestFixture]
    public class MatchResultTests
    {

        /// <summary>
        /// Tests .Equals method with an object of another type
        /// </summary>
        [Test]
        public void EqualsTestAnotherObject()
        {
            MatchResult mr1 = new MatchResult();
            mr1.StartIndex = 1;
            mr1.EndIndex = 2;
            mr1.Result = true;
            Assert.IsFalse(mr1.Equals("Test"));
        }

        /// <summary>
        /// Tests .Equals method with a different MatchResult object
        /// </summary>
        [Test]
        public void EqualsTestDifferentMatchResult()
        {
            MatchResult mr1 = new MatchResult();
            mr1.StartIndex = 1;
            mr1.EndIndex = 2;
            mr1.Result = true;
            MatchResult mr2 = new MatchResult();
            mr2.StartIndex = 2;
            mr2.EndIndex = 2;
            mr2.Result = true;
            Assert.IsFalse(mr1.Equals(mr2));
        }

        /// <summary>
        /// Tests .Equals method with the same MatchResult object
        /// </summary>
        [Test]
        public void EqualsTestSameMatchResult()
        {
            MatchResult mr1 = new MatchResult();
            mr1.StartIndex = 1;
            mr1.EndIndex = 2;
            mr1.Result = true;
            MatchResult mr2 = new MatchResult();
            mr2.StartIndex = 1;
            mr2.EndIndex = 2;
            mr2.Result = true;
            Assert.IsTrue(mr1.Equals(mr2));
        }

        /// <summary>
        /// Tests == operator
        /// </summary>
        [Test]
        public void EqualOperatorTest()
        {
            MatchResult mr1 = new MatchResult();
            mr1.StartIndex = 1;
            mr1.EndIndex = 2;
            mr1.Result = true;
            MatchResult mr2 = new MatchResult();
            mr2.StartIndex = 2;
            mr2.EndIndex = 2;
            mr2.Result = true;
            Assert.IsFalse(mr1 == mr2, "Equality operator fails when objects are different");
            mr2.StartIndex = 1;
            Assert.IsTrue(mr1 == mr2, "Equality operator fails when objects are same");
        }

        /// <summary>
        /// Tests != operator
        /// </summary>
        [Test]
        public void NotEqualOperatorTest()
        {
            MatchResult mr1 = new MatchResult();
            mr1.StartIndex = 1;
            mr1.EndIndex = 2;
            mr1.Result = true;
            MatchResult mr2 = new MatchResult();
            mr2.StartIndex = 2;
            mr2.EndIndex = 2;
            mr2.Result = true;
            Assert.IsTrue(mr1 != mr2, "Inequality operator fails when objects are different");
            mr2.StartIndex = 1;
            Assert.IsFalse(mr1 != mr2, "Inequality operator fails when objects are same");
        }

    }
}
