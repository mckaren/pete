using System;
using System.Globalization;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class SkipCommandTests
    {
        [TestMethod]
        public void ParseArgumentsSetsRowsToSkip()
        {
            SkipCommand cmd = new SkipCommand();
            cmd.ArgumentsString = @"15";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual(15, cmd.RowsToSkip);
        }

        [TestMethod]
        public void ApplyToDataRemovesRowsFromTheData()
        {
            SkipCommand cmd = new SkipCommand() { RowsToSkip = 1 };
            DataElement da = Helpers.CreateTestDataElement();
            cmd.ApplyToData(da);
            Assert.AreEqual(1, da.Rows.Count);
            Assert.AreEqual(1, da.Columns[0].Data.Count);
            Assert.AreEqual(1, da.Columns[1].Data.Count);
            Assert.AreEqual(1, da.Columns[2].Data.Count);
            Assert.AreEqual(5.05, da.Columns[0].Data[0]);
            Assert.AreEqual("client 2", da.Columns[1].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[0]);
            Assert.AreEqual(da.Columns[0].Data[0], da.Rows[0].Data[0]);
            Assert.AreEqual(da.Columns[1].Data[0], da.Rows[0].Data[1]);
            Assert.AreEqual(da.Columns[2].Data[0], da.Rows[0].Data[2]);
        }
    }
}
