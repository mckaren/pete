using System;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using McKinsey.PowerPointGenerator.Extensions;

namespace McKinsey.PowerPointGenerator.Tests.Extensions
{
    [TestClass]
    public class RunExtensionsTests
    {
        [TestMethod]
        public void ReplaceReplacesTextInTheRunWithoutFill()
        {
            Run p = new Run(@"<a:r xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:rPr lang=""de-DE"" dirty=""0"" smtClean=""0"" /><a:t>here goes #value[0][1]:{0.0000}# so we hope</a:t></a:r>");
            p.Replace("#value[0][1]:{0.0000}#", "5.0500");
            Assert.AreEqual("here goes 5.0500 so we hope", p.InnerText);
        }
    }
}
