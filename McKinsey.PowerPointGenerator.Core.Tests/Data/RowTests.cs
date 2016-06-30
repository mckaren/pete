using System;
using System.Linq;
using McKinsey.PowerPointGenerator.Core.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Core.Tests.Data
{
    [TestClass]
    public class RowTests
    {
        [TestMethod]
        public void GetHeaderReturnsMappedHeaderIfSet()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            Assert.AreEqual("row 2", row.GetHeader());
            row.MappedHeader = "client";
            Assert.AreEqual("client", row.GetHeader());

        }
        [TestMethod]
        public void IndexByIntReturnsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            Assert.AreEqual(5.05, row[0]);
            Assert.AreEqual("client 2", row[1]);
            Assert.AreEqual(true, row[2]);
        }

        [TestMethod]
        public void IndexByNameReturnsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            Assert.AreEqual(5.05, row["Column 1"]);
            Assert.AreEqual("client 2", row["Column 2"]);
            Assert.AreEqual(true, row["Column 3"]);
        }

        [TestMethod]
        public void IndexByIndexReturnsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            Assert.AreEqual(5.05, row[new Index("0")]);
            Assert.AreEqual("client 2", row[new Index("1")]);
            Assert.AreEqual(true, row[new Index("2")]);
            Assert.AreEqual(5.05, row[new Index("Column 1")]);
            Assert.AreEqual("client 2", row[new Index("Column 2")]);
            Assert.AreEqual(true, row[new Index("Column 3")]);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void IndexByIntThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            var test = row[4];
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void IndexByNameThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            var test =  row["My Column 1"];
        }

        [TestMethod]
        public void IndexByIndexSetsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            row[new Index("0")] = 10.25;
            row[new Index("1")] = "Client 5";
            row[new Index("2")] = false;
            Assert.AreEqual(10.25, row.Data[0]);
            Assert.AreEqual("Client 5", row.Data[1]);
            Assert.AreEqual(false, row.Data[2]);
            row[new Index("Column 1")] = 20.25;
            row[new Index("Column 2")] = "Client 3";
            row[new Index("Column 3")] = true;
            Assert.AreEqual(20.25, row.Data[0]);
            Assert.AreEqual("Client 3", row.Data[1]);
            Assert.AreEqual(true, row.Data[2]);
        }

        [TestMethod]
        public void IndexByIntSetsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            row[0] = 10.25;
            row[1] = "Client 5";
            row[2] = false;
            Assert.AreEqual(10.25, row.Data[0]);
            Assert.AreEqual("Client 5", row.Data[1]);
            Assert.AreEqual(false, row.Data[2]);
        }

        [TestMethod]
        public void IndexByNameSetsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            row["Column 1"] = 10.25;
            row["Column 2"] = "Client 5";
            row["Column 3"] = false;
            Assert.AreEqual(10.25, row.Data[0]);
            Assert.AreEqual("Client 5", row.Data[1]);
            Assert.AreEqual(false, row.Data[2]);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void IndexByIntOnSetThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            row[4] = 1;
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void IndexByNameOnSetThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Rows[1];
            row["My Column 1"] = "test";
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void IndexByNameThrowsArgumentExceptionWhenNoRowHeaders()
        {
            DataElement da = Helpers.CreateTestDataElement();
            da.HasColumnHeaders = false;
            da.HasRowHeaders = false;
            var row = da.Rows[1];
            var test = row["Column 1"];
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void IndexByNameOnSetThrowsArgumentExceptionWhenNoRowHeaders()
        {
            DataElement da = Helpers.CreateTestDataElement();
            da.HasColumnHeaders = false;
            da.HasRowHeaders = false;
            var row = da.Rows[1];
            row["Column 1"] = "test";
        }
    }
}
