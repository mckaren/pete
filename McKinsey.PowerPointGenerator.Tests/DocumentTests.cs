using System;
using System.Linq;
using System.IO;
using McKinsey.PowerPointGenerator.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests
{
    [TestClass]
    public class DocumentTests
    {
        [TestMethod]
        public void LoadLoadsPowerpointFile()
        {
            Document doc = new Document();
            var s = Helpers.GetTemplateStreamFromResources(Resources.ThreeSlides);
            doc.Load(s);
            Assert.IsNotNull(doc.PptDocument);
            Assert.IsNotNull(doc.PresentationPart);
        }

        [TestMethod]
        public void GetSlidesGetsAllNotHiddenSlides()
        {
            Document doc = new Document();
            var s = Helpers.GetTemplateStreamFromResources(Resources.ThreeSlides);
            doc.Load(s);
            doc.GetSlides();
            Assert.AreEqual(2, doc.Slides.Count()); //3 but one is hidden
            Assert.AreEqual(0, doc.Slides[0].Number);
            Assert.AreEqual(2, doc.Slides[1].Number);
        }
    }
}
