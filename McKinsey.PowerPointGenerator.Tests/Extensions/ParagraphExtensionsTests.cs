using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using McKinsey.PowerPointGenerator.Extensions;

namespace McKinsey.PowerPointGenerator.Tests.Extensions
{
    [TestClass]
    public class ParagraphExtensionsTests
    {
        [TestMethod]
        public void ReplaceReplacesTextInTheParagraphThatIsSpreadIntoManyRuns()
        {
            Paragraph p = new Paragraph(@"<a:p xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:r><a:rPr lang=""de-DE"" dirty=""0"" /><a:t>#</a:t></a:r><a:r><a:rPr lang=""de-DE"" dirty=""0"" smtClean=""0"" /><a:t>Client1Name#</a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t>’s total IT spend is in line with that of peers, coming in between the median and </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" smtClean=""0"" /><a:t>top </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t>quartile marks</a:t></a:r><a:endParaRPr lang=""de-CH"" dirty=""0"" /></a:p>");
            p.Replace("#Client1Name#", "IBM");
            Assert.AreEqual("IBM’s total IT spend is in line with that of peers, coming in between the median and top quartile marks", p.InnerText);
        }

        [TestMethod]
        public void ReplaceMergesRequiredRuns()
        {
            Paragraph p = new Paragraph(@"<a:p xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:r><a:rPr lang=""de-DE"" dirty=""0"" /><a:t>#</a:t></a:r><a:r><a:rPr lang=""de-DE"" dirty=""0"" smtClean=""0"" /><a:t>Client1Name#</a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t>’s total IT spend is in line with that of peers, coming in between the median and </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" smtClean=""0"" /><a:t>top </a:t></a:r><a:r><a:rPr lang=""en-US"" dirty=""0"" /><a:t>quartile marks</a:t></a:r><a:endParaRPr lang=""de-CH"" dirty=""0"" /></a:p>");
            Assert.AreEqual(5, p.Elements<Run>().Count());
            Assert.AreEqual("#", p.Elements<Run>().ElementAt(0).InnerText);
            Assert.AreEqual("Client1Name#", p.Elements<Run>().ElementAt(1).InnerText);
            Assert.AreEqual("’s total IT spend is in line with that of peers, coming in between the median and ", p.Elements<Run>().ElementAt(2).InnerText);
            Assert.AreEqual("top ", p.Elements<Run>().ElementAt(3).InnerText);
            Assert.AreEqual("quartile marks", p.Elements<Run>().ElementAt(4).InnerText);

            p.Replace("#Client1Name#", "IBM");
            Assert.AreEqual(4, p.Elements<Run>().Count());
            Assert.AreEqual("IBM", p.Elements<Run>().ElementAt(0).InnerText);
            Assert.AreEqual("’s total IT spend is in line with that of peers, coming in between the median and ", p.Elements<Run>().ElementAt(1).InnerText);
            Assert.AreEqual("top ", p.Elements<Run>().ElementAt(2).InnerText);
            Assert.AreEqual("quartile marks", p.Elements<Run>().ElementAt(3).InnerText);
        }
    }
}
