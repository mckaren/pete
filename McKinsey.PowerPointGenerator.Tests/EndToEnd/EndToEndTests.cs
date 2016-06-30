using System;
using System.IO;
using System.Linq;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Processing;
using McKinsey.PowerPointGenerator.Tests.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace McKinsey.PowerPointGenerator.Tests
{
    [TestClass]
    public class EndToEndTests
    {
        [TestMethod]
        [Ignore]
        public void ParseGetsAllElementsWithNamesAndReferencesToXmlElements()
        {
            Document doc = new Document();
            string path = @"c:\test.pptx";
            File.WriteAllBytes(path, Resources.ThreeSlides);
            //var s = Helpers.GetTemplateStreamFromResources(Resources.ThreeSlides);
            var s = File.Open(path, FileMode.Open, FileAccess.ReadWrite);
            doc.Load(s);
            doc.GetSlides();
            doc.Slides[0].DiscoverShapes();
            doc.Slides[0].DiscoverCommands();
            var chartElement = doc.Slides[0].Shapes.First(c => c is TableElement);
            DataElement da = Helpers.CreateSingleValueElement("customer", "IBM");
            da.HasRowHeaders = true;
            da.HasColumnHeaders = true;
            chartElement.Data = da;
            TableElementProcessor proc = new TableElementProcessor();
            proc.Process(chartElement);
            doc.Slides[0].Slide.Save();
            doc.SaveAndClose();
        }
    }
}
