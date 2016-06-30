using System;
using System.Globalization;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class SortCommandTests
    {
        [TestMethod]
        public void ParseArgumentsSetsIndexAndAscendingWhenNoOrderProvided()
        {
            SortCommand cmd = new SortCommand();
            cmd.ArgumentsString = @"""Column 2""";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual("Column 2", cmd.Index.Name);
            Assert.AreEqual(SortOrder.Ascending, cmd.SortOrder);
        }

        [TestMethod]
        public void ParseArgumentsSetsIndexAndAscendingWhenAscendingProvided()
        {
            SortCommand cmd = new SortCommand();
            cmd.ArgumentsString = @"2 ASC";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual(2, cmd.Index.Number);
            Assert.AreEqual(SortOrder.Ascending, cmd.SortOrder);
        }

        [TestMethod]
        public void ParseArgumentsSetsIndexAndDescendingWhenDescendingProvided()
        {
            SortCommand cmd = new SortCommand();
            cmd.ArgumentsString = @"""col 3"" DESC";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual("col 3", cmd.Index.Name);
            Assert.AreEqual(SortOrder.Descending, cmd.SortOrder);
        }

        [TestMethod]
        public void ApplyToDataSortsData()
        {
            SortCommand cmd = new SortCommand() { Index = new Index("column 1"), SortOrder = SortOrder.Descending  };
            DataElement da = Helpers.CreateTestDataElement();
            cmd.ApplyToData(da);
            Assert.AreEqual(5.05, da.Columns[0].Data[0]);
            Assert.AreEqual(1.0, da.Columns[0].Data[1]);
            Assert.AreEqual("client 2", da.Columns[1].Data[0]);
            Assert.AreEqual("client 1", da.Columns[1].Data[1]);
            Assert.AreEqual(true, da.Columns[2].Data[0]);
            Assert.AreEqual(false, da.Columns[2].Data[1]);
            Assert.AreEqual(da.Columns[0].Data[0], da.Rows[0].Data[0]);
            Assert.AreEqual(da.Columns[0].Data[1], da.Rows[1].Data[0]);
            Assert.AreEqual(da.Columns[1].Data[0], da.Rows[0].Data[1]);
            Assert.AreEqual(da.Columns[1].Data[1], da.Rows[1].Data[1]);
            Assert.AreEqual(da.Columns[2].Data[0], da.Rows[0].Data[2]);
            Assert.AreEqual(da.Columns[2].Data[1], da.Rows[1].Data[2]);
        }
    }
}
