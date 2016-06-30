using System;
using McKinsey.PowerPointGenerator.Core.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Core.Tests.Data
{
    [TestClass]
    public class ColumnTests
    {
        [TestMethod]
        public void GetHeaderReturnsMappedHeaderIfSet()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            Assert.AreEqual("Column 2", column.GetHeader());
            column.MappedHeader = "client";
            Assert.AreEqual("client", column.GetHeader());
        }

        [TestMethod]
        public void IndexByIndexReturnsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            Assert.AreEqual("client 1", column[new Index("0")]);
            Assert.AreEqual("client 2", column[new Index("1")]);
            Assert.AreEqual("client 1", column[new Index("row 1")]);
            Assert.AreEqual("client 2", column[new Index("row 2")]);
        }

        [TestMethod]
        public void IndexByIntReturnsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            Assert.AreEqual("client 1", column[0]);
            Assert.AreEqual("client 2", column[1]);
        }

        [TestMethod]
        public void IndexByNameReturnsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            Assert.AreEqual("client 1", column["row 1"]);
            Assert.AreEqual("client 2", column["row 2"]);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void IndexByIntThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            var test = column[4];
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void IndexByNameThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            var test = column["My row 1"];
        }

        [TestMethod]
        public void IndexByIndexSetsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            column[new Index("0")] = "Client 4";
            column[new Index("1")] = "Client 5";
            Assert.AreEqual("Client 4", column.Data[0]);
            Assert.AreEqual("Client 5", column.Data[1]);
            column[new Index("row 1")] = "Client 6";
            column[new Index("row 2")] = "Client 7";
            Assert.AreEqual("Client 6", column.Data[0]);
            Assert.AreEqual("Client 7", column.Data[1]);
        }

        [TestMethod]
        public void IndexByIntSetsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            column[0] = "Client 4";
            column[1] = "Client 5";
            Assert.AreEqual("Client 4", column.Data[0]);
            Assert.AreEqual("Client 5", column.Data[1]);
        }

        [TestMethod]
        public void IndexByNameSetsData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            column["row 1"] = "Client 4";
            column["row 2"] = "Client 5";
            Assert.AreEqual("Client 4", column.Data[0]);
            Assert.AreEqual("Client 5", column.Data[1]);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void IndexByIntOnSetThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            column[4] = 1;
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void IndexByNameOnSetThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Columns[1];
            column["My row 1"] = "test";
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void IndexByNameThrowsArgumentExceptionWhenNoRowHeaders()
        {
            DataElement da = Helpers.CreateTestDataElement();
            da.HasColumnHeaders = false;
            da.HasRowHeaders = false;
            var column = da.Columns[1];
            var test = column["Column 1"];
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void IndexByNameOnSetThrowsArgumentExceptionWhenNoRowHeaders()
        {
            DataElement da = Helpers.CreateTestDataElement();
            da.HasColumnHeaders = false;
            da.HasRowHeaders = false;
            var column = da.Columns[1];
            column["row 1"] = "test";
        }
    }
}
