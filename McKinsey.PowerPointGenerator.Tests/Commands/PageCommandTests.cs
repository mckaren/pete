using System;
using System.Globalization;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class PageCommandTests
    {
        [TestMethod]
        public void ParseArgumentsSetsRowsAndMaximumColumnsPerPageWhenOnlyRowsSpecified()
        {
            PageCommand cmd = new PageCommand();
            cmd.ArgumentsString = @"10";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual(10, cmd.RowsPerPage);
            Assert.AreEqual(Int32.MaxValue, cmd.ColumnsPerPage);
        }

        [TestMethod]
        public void ParseArgumentsSetsRowsAndColumnsPerPageWhenBothSpecified()
        {
            PageCommand cmd = new PageCommand();
            cmd.ArgumentsString = @"15, 5";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual(15, cmd.RowsPerPage);
            Assert.AreEqual(5, cmd.ColumnsPerPage);
        }

        [TestMethod]
        public void ParseArgumentsSetsMaximumRowsAndColumnsPerPageWhenEmptyArguments()
        {
            PageCommand cmd = new PageCommand();
            cmd.ArgumentsString = @"";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual(Int32.MaxValue, cmd.RowsPerPage);
            Assert.AreEqual(Int32.MaxValue, cmd.ColumnsPerPage);
        }
    }
}
