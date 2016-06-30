using System;
using System.Globalization;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class TakeCommandTests
    {
        [TestMethod]
        public void ParseArgumentsSetsRowsToTake()
        {
            TakeCommand cmd = new TakeCommand();
            cmd.ArgumentsString = @"10";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual(10, cmd.RowsToTake);
        }

        [TestMethod]
        public void ApplyToDataRemovesRemainignRows()
        {
            TakeCommand cmd = new TakeCommand() { RowsToTake = 1 };
            DataElement da = Helpers.CreateTestDataElement();
            cmd.ApplyToData(da);
            Assert.AreEqual(1, da.Rows.Count);
            Assert.AreEqual(1, da.Columns[0].Data.Count);
            Assert.AreEqual(1, da.Columns[1].Data.Count);
            Assert.AreEqual(1, da.Columns[2].Data.Count);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual("client 1", da.Columns[1].Data[0]);
            Assert.AreEqual(false, da.Columns[2].Data[0]);
            Assert.AreEqual(da.Columns[0].Data[0], da.Rows[0].Data[0]);
            Assert.AreEqual(da.Columns[1].Data[0], da.Rows[0].Data[1]);
            Assert.AreEqual(da.Columns[2].Data[0], da.Rows[0].Data[2]);
        }
    }
}
