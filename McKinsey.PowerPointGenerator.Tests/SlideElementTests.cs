using System;
using System.Collections.Generic;
using System.Linq;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace McKinsey.PowerPointGenerator.Tests
{
    [TestClass]
    public class SlideElementTests
    {
        [TestMethod]
        public void ParseGetsAllElementsWithNamesAndReferencesToXmlElements()
        {
            Document doc = new Document();
            var s = Helpers.GetTemplateStreamFromResources(Resources.ThreeSlides);
            doc.Load(s);
            doc.GetSlides();
            doc.Slides[1].DiscoverShapes();
            Assert.AreEqual(4, doc.Slides[1].Shapes.Count);
            Assert.AreEqual(1, doc.Slides[1].Shapes.Count(sh => sh is TableElement));
            Assert.AreEqual(1, doc.Slides[1].Shapes.Count(sh => sh is ChartElement));
            Assert.AreEqual(1, doc.Slides[1].Shapes.Count(sh => sh is ShapeElement));
            Assert.AreEqual(1, doc.Slides[1].Shapes.Count(sh => sh is TextElement));
            var table = doc.Slides[1].Shapes.First(sh => sh is TableElement);
            var chart = doc.Slides[1].Shapes.First(sh => sh is ChartElement);
            var shape = doc.Slides[1].Shapes.First(sh => sh is ShapeElement);
            var text = doc.Slides[1].Shapes.First(sh => sh is TextElement);
            Assert.AreEqual("SpendByRevenuesA", table.Name);
            Assert.AreEqual("SpendByRevenuesC", chart.Name);
            Assert.AreEqual("SpendByRevenuesB", shape.Name);
            Assert.AreEqual("Client1Name", text.Name);
        }
    }
}
