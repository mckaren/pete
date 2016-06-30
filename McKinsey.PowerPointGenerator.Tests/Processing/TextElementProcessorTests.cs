using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Processing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Processing
{
    [TestClass]
    public class TextElementProcessorTests
    {
        [TestMethod]
        public void ProcessReplacesTagInTheElement()
        {
            Paragraph p = new Paragraph(@"<a:p xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:r><a:rPr lang=""de-DE"" dirty=""0"" /><a:t>#</a:t></a:r><a:r><a:rPr lang=""de-DE"" dirty=""0"" smtClean=""0"" /><a:t>Client1Name#</a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t>’s total IT spend is in line with that of peers, coming in between the median and </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" smtClean=""0"" /><a:t>top </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t>quartile marks</a:t></a:r><a:endParaRPr lang=""de-CH"" dirty=""0"" /></a:p>");
            TextElementProcessor processor = new TextElementProcessor();
            TextElement element = TextElement.Create("Client1Name", p, null);
            element.Data = Helpers.CreateSingleValueElement("client1Name", "IBM");
            processor.Process(element);
            Assert.AreEqual("IBM’s total IT spend is in line with that of peers, coming in between the median and top quartile marks", p.InnerText);
        }

        [TestMethod]
        public void ProcessReplacesTagWithIndexesInTheElement()
        {
            Paragraph p = new Paragraph(@"<a:p xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:r><a:rPr lang=""de-DE"" dirty=""0"" /><a:t>#</a:t></a:r><a:r><a:rPr lang=""de-DE"" dirty=""0"" smtClean=""0"" /><a:t>Client1Name[""column 2""][""row 1""]#</a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t>’s total IT spend is in line with that of peers, coming in between the median and </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" smtClean=""0"" /><a:t>top </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t>quartile marks</a:t></a:r><a:endParaRPr lang=""de-CH"" dirty=""0"" /></a:p>");
            TextElementProcessor processor = new TextElementProcessor();
            TextElement element = TextElement.Create(@"Client1Name[""column 2""][""row 1""]", p, null);
            element.Data = Helpers.CreateTestDataElement();
            element.ColumnIndexes.Add(new Core.Data.Index("column 2"));
            element.RowIndexes.Add(new Core.Data.Index("row 1"));
            processor.Process(element);
            Assert.AreEqual("client 1’s total IT spend is in line with that of peers, coming in between the median and top quartile marks", p.InnerText);
        }

        [TestMethod]
        public void ProcessReplacesFormattedTagInTheElement()
        {
            Paragraph p = new Paragraph(@"<a:p xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:r><a:rPr lang=""de-DE"" dirty=""0"" /><a:t>#</a:t></a:r><a:r><a:rPr lang=""de-DE"" dirty=""0"" smtClean=""0"" /><a:t>salesTotal</a:t></a:r><a:r><a:rPr lang=""de-DE"" dirty=""0"" smtClean=""0"" /><a:t>:(“##,#”)#</a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t> total IT spend is in line with that of peers, coming in between the median and </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" smtClean=""0"" /><a:t>top </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t>quartile marks</a:t></a:r><a:endParaRPr lang=""de-CH"" dirty=""0"" /></a:p>");
            TextElementProcessor processor = new TextElementProcessor();
            TextElement element = TextElement.Create("salesTotal:(“##,#”)", p, null);
            element.Data = Helpers.CreateSingleValueElement("salesTotal", 2150247);
            element.DiscoverCommands();
            processor.Process(element);
            Assert.AreEqual("2,150,247 total IT spend is in line with that of peers, coming in between the median and top quartile marks", p.InnerText);
        }    }
}
