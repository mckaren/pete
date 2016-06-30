using System;
using System.Collections.Generic;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Elements
{
    [TestClass]
    public class IndexTests
    {
        [TestMethod]
        public void ConstructorSetsNumberOnIntString()
        {
            var index = new Index("34");
            Assert.AreEqual(34, index.Number);
        }

        [TestMethod]
        public void ConstructorSetsNameWhenStringValue()
        {
            var index = new Index("Column 2");
            Assert.AreEqual("Column 2", index.Name);
            Assert.IsNull(index.Number);
        }
    }
}
