using System;
using DocumentFormat.OpenXml.Presentation;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Processing;
using McKinsey.PowerPointGenerator.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Processing
{
    [TestClass]
    public class ShapeElementProcessorTests
    {
        [TestMethod]
        public void ProcessReplacesWholeShapeContent()
        {
            Shape s = new Shape(Resources.TestShape);
            ShapeElementProcessor processor = new ShapeElementProcessor();
            ShapeElement element = ShapeElement.Create("customer", s, null);
            element.Data = Helpers.CreateSingleValueElement("customer", "IBM");
            processor.Process(element);
            Assert.AreEqual("IBM", s.InnerText);
        }

        [TestMethod]
        public void ProcessReplacesFormattedTagInTheElement()
        {
            Shape s = new Shape(Resources.TestShapeWithTag);
            ShapeElementProcessor processor = new ShapeElementProcessor();
            ShapeElement element = ShapeElement.Create("customer", s, null);
            element.Data = Helpers.CreateSingleValueElement("customer", "IBM");
            processor.Process(element);
            Assert.AreEqual("Business profile, IBM, 2014", s.InnerText);
        }
    }
}
