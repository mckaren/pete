using System;
using System.Globalization;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class TransposeCommandTests
    {
        [TestMethod]
        public void ApplyToDataTransposesTheData()
        {
            TransposeCommand cmd = new TransposeCommand();
            DataElement da = Helpers.CreateTestDataElement();
            cmd.ApplyToData(da);
            Assert.AreEqual(1.0, da.Rows[0].Data[0]);
            Assert.AreEqual(5.05, da.Rows[0].Data[1]);
            Assert.AreEqual("client 1", da.Rows[1].Data[0]);
            Assert.AreEqual("client 2", da.Rows[1].Data[1]);
            Assert.AreEqual(false, da.Rows[2].Data[0]);
            Assert.AreEqual(true, da.Rows[2].Data[1]);
            Assert.AreEqual("Column 1", da.Rows[0].Header);
            Assert.AreEqual("Column 2", da.Rows[1].Header);
            Assert.AreEqual("IBM", da.Rows[1].MappedHeader);
            Assert.AreEqual("Column 3", da.Rows[2].Header);
            Assert.IsTrue(da.Rows[2].IsHidden);
            Assert.AreEqual("row 1", da.Columns[0].Header);
            Assert.AreEqual("Totals", da.Columns[0].MappedHeader);
            Assert.AreEqual("row 2", da.Columns[1].Header);
            Assert.IsTrue(da.Columns[1].IsHidden);
            Assert.AreEqual(da.Rows[0].Data[0], da.Columns[0].Data[0]);
            Assert.AreEqual(da.Rows[0].Data[1], da.Columns[1].Data[0]);
            Assert.AreEqual(da.Rows[1].Data[0], da.Columns[0].Data[1]);
            Assert.AreEqual(da.Rows[1].Data[1], da.Columns[1].Data[1]);
            Assert.AreEqual(da.Rows[2].Data[0], da.Columns[0].Data[2]);
            Assert.AreEqual(da.Rows[2].Data[1], da.Columns[1].Data[2]);
        }
    }
}
