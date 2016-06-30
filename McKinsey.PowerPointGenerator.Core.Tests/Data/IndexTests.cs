using System;
using McKinsey.PowerPointGenerator.Core.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Core.Tests.Data
{
    [TestClass]
    public class IndexTests
    {
        [TestMethod]
        public void COnstructorWithStarCreatesAll()
        {
        	Assert.IsTrue((new Index("*")).IsAll);
        	Assert.IsFalse((new Index("column")).IsAll);
        }

        [TestMethod]
        public void EqualsReturnsCorrectResults()
        {
            Assert.IsTrue((new Index(5)) == (new Index(5)));
            Assert.IsFalse((new Index(5)) == (new Index(6)));
            Assert.IsTrue((new Index("column 1")) == (new Index("ColumN 1")));
            Assert.IsFalse((new Index("column 1")) == (new Index("column 2")));
            Assert.IsTrue((new Index("*")) == (new Index("*")));
            Assert.IsFalse((new Index("*")) == (new Index(1)));
            Assert.IsFalse((new Index("column 1")) == (new Index(1)));
            Assert.IsFalse((new Index("*")) == (new Index("column")));
            Assert.IsFalse((new Index(1)) == null);
            var idx = new Index(1);
            Assert.IsTrue(idx == idx);
        }
    }
}
