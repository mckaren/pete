using System;
using System.Collections.Generic;
using System.Linq;
using McKinsey.PowerPointGenerator.Core.Data;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace McKinsey.PowerPointGenerator.Core.Tests.Data
{
    [TestClass]
    public class DataElementTests
    {
        [TestMethod]
        public void DataElementDeserializesFromJson()
        {
            string json = Helpers.CreateJsonForDataElementWithDataInColumns();
            DataElement da = JsonConvert.DeserializeObject<DataElement>(json);
            Assert.ReferenceEquals(da, da.Rows[0].ParentElement);
            Assert.ReferenceEquals(da, da.Rows[1].ParentElement);
            Assert.ReferenceEquals(da, da.Columns[0].ParentElement);
            Assert.ReferenceEquals(da, da.Columns[1].ParentElement);
            Assert.ReferenceEquals(da, da.Columns[2].ParentElement);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual(5.05, da.Columns[0].Data[1]);
            Assert.AreEqual("client 1", da.Columns[1].Data[0]);
            Assert.AreEqual("client 2", da.Columns[1].Data[1]);
            Assert.AreEqual(false, da.Columns[2].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[1]);
            Assert.AreEqual(da.Columns[0].Data[0], da.Rows[0].Data[0]);
            Assert.AreEqual(da.Columns[0].Data[1], da.Rows[1].Data[0]);
            Assert.AreEqual(da.Columns[1].Data[0], da.Rows[0].Data[1]);
            Assert.AreEqual(da.Columns[1].Data[1], da.Rows[1].Data[1]);
            Assert.AreEqual(da.Columns[2].Data[0], da.Rows[0].Data[2]);
            Assert.AreEqual(da.Columns[2].Data[1], da.Rows[1].Data[2]);
        }

        [TestMethod]
        public void NormalizeSetsParentElementAndDataInColumns()
        {
            DataElement da = Helpers.CreateTestDataElement();
            Assert.ReferenceEquals(da, da.Rows[0].ParentElement);
            Assert.ReferenceEquals(da, da.Rows[1].ParentElement);
            Assert.ReferenceEquals(da, da.Columns[0].ParentElement);
            Assert.ReferenceEquals(da, da.Columns[1].ParentElement);
            Assert.ReferenceEquals(da, da.Columns[2].ParentElement);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual(5.05, da.Columns[0].Data[1]);
            Assert.AreEqual("client 1", da.Columns[1].Data[0]);
            Assert.AreEqual("client 2", da.Columns[1].Data[1]);
            Assert.AreEqual(false, da.Columns[2].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[1]);
            Assert.AreEqual(da.Columns[0].Data[0], da.Rows[0].Data[0]);
            Assert.AreEqual(da.Columns[0].Data[1], da.Rows[1].Data[0]);
            Assert.AreEqual(da.Columns[1].Data[0], da.Rows[0].Data[1]);
            Assert.AreEqual(da.Columns[1].Data[1], da.Rows[1].Data[1]);
            Assert.AreEqual(da.Columns[2].Data[0], da.Rows[0].Data[2]);
            Assert.AreEqual(da.Columns[2].Data[1], da.Rows[1].Data[2]);
        }

        [TestMethod]
        public void NormalizeAlsoSetsLegendsIfTheyExists()
        {
            DataElement da = Helpers.CreateNotNormalisedDataAleemntWithDataInColumns();
            da.Columns[1].Legends.Add(null);
            da.Columns[1].Legends.Add("legend 1");
            da.Normalize();
            Assert.IsNull(da.Rows[0].Legends[1]);
            Assert.AreEqual("legend 1", da.Rows[1].Legends[1]);
        }

        [TestMethod]
        public void NormalizeSetsMissingColumns()
        {
            DataElement da = new DataElement();
            da.Rows.Add(new Row());
            da.Rows[0].Data.Add(5);
            da.Rows[0].Data.Add(4);
            da.Rows[0].Data.Add(3);
            da.Rows[0].Data.Add(2);
            da.Rows[0].Data.Add(1);
            da.Normalize();
            Assert.AreEqual(5, da.Columns.Count);
        }

        [TestMethod]
        public void GetFragmentByIndexesReturnsFragmentByIndexes()
        {
            DataElement da = Helpers.CreateTestDataElementWithTwoNumericColumns();
            List<Index> rows = new List<Index> { new Index(1) };
            List<Index> cols = new List<Index> { new Index(0), new Index(3) };
            var frag = da.GetFragmentByIndexes(rows, cols);
            Assert.AreEqual(5.05, frag.Rows[0].Data[0]);
            Assert.AreEqual(4.05, frag.Rows[0].Data[1]);
            Assert.AreEqual(1, frag.Rows.Count);
            Assert.AreEqual(2, frag.Columns.Count);
            Assert.AreEqual(2, frag.Rows[0].Data.Count);
            Assert.AreEqual(2, frag.Rows[0].Legends.Count);
            Assert.AreEqual(1, frag.Columns[0].Data.Count);
            Assert.AreEqual(1, frag.Columns[0].Legends.Count);
            Assert.AreEqual(1, frag.Columns[1].Data.Count);
            Assert.AreEqual(1, frag.Columns[1].Legends.Count);
        }

        [TestMethod]
        public void GetFragmentByIndexesReturnsFragmentByIndexesAndSetsIsHiddenColumnsAndRows()
        {
            DataElement da = Helpers.CreateTestDataElementWithTwoNumericColumns();
            List<Index> rows = new List<Index> { new Index(0) { IsHidden = true }, new Index(1) };
            List<Index> cols = new List<Index> { new Index(0), new Index(1) { IsHidden = true }, new Index(3) };
            var frag = da.GetFragmentByIndexes(rows, cols);
            Assert.AreEqual(2, frag.Rows.Count);
            Assert.AreEqual(3, frag.Columns.Count);
            Assert.IsTrue(frag.Rows[0].IsHidden);
            Assert.IsFalse(frag.Rows[1].IsHidden);
            Assert.IsFalse(frag.Columns[0].IsHidden);
            Assert.IsTrue(frag.Columns[1].IsHidden);
            Assert.IsFalse(frag.Columns[2].IsHidden);
        }

        [TestMethod]
        public void GetFragmentByIndexesReturnsTheSameElementWhenNoIndexes()
        {
            DataElement da1 = Helpers.CreateTestDataElement();
            List<Index> rows = new List<Index>();
            List<Index> cols = new List<Index>();
            var da = da1.GetFragmentByIndexes(rows, cols);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual(5.05, da.Columns[0].Data[1]);
            Assert.AreEqual("client 1", da.Columns[1].Data[0]);
            Assert.AreEqual("client 2", da.Columns[1].Data[1]);
            Assert.AreEqual(false, da.Columns[2].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[1]);
        }

        [TestMethod]
        public void GetFragmentByIndexesReturnsTheSameElementWhenNoRowIndexes()
        {
            DataElement da1 = Helpers.CreateTestDataElement();
            List<Index> rows = new List<Index>();
            List<Index> cols = new List<Index> { new Index(1) };
            var da = da1.GetFragmentByIndexes(rows, cols);
            Assert.AreEqual("client 1", da.Columns[0].Data[0]);
            Assert.AreEqual("client 2", da.Columns[0].Data[1]);
        }

        [TestMethod]
        public void GetFragmentByIndexesReturnsTheSameElementWhenNoColumnIndexes()
        {
            DataElement da1 = Helpers.CreateTestDataElement();
            List<Index> rows = new List<Index> { new Index(1) };
            List<Index> cols = new List<Index>();
            var da = da1.GetFragmentByIndexes(rows, cols);
            Assert.AreEqual(5.05, da.Columns[0].Data[0]);
            Assert.AreEqual("client 2", da.Columns[1].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[0]);
        }

        [TestMethod]
        public void DataReturnsAndSetsSingleCell()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var data = da.Data(new Index(1), new Index(1));
            Assert.AreEqual("client 2", data);
            da.Data(new Index(1), new Index(1), "test");
            Assert.AreEqual("test", da.Data(new Index(1), new Index(1)));
        }

        [TestMethod]
        public void RowByIntReturnsRow()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Row(1);
            Assert.AreEqual(5.05, row.Data[0]);
            Assert.AreEqual("client 2", row.Data[1]);
            Assert.AreEqual(true, row.Data[2]);
        }

        [TestMethod]
        public void RowByNameReturnsRow()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Row("row 2");
            Assert.AreEqual(5.05, row.Data[0]);
            Assert.AreEqual("client 2", row.Data[1]);
            Assert.AreEqual(true, row.Data[2]);
        }

        [TestMethod]
        public void RowByIndexRow()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Row(new Index("1"));
            Assert.AreEqual(5.05, row.Data[0]);
            Assert.AreEqual("client 2", row.Data[1]);
            Assert.AreEqual(true, row.Data[2]);
            row = da.Row(new Index("row 2"));
            Assert.AreEqual(5.05, row.Data[0]);
            Assert.AreEqual("client 2", row.Data[1]);
            Assert.AreEqual(true, row.Data[2]);
        }

        [TestMethod]
        public void ColumnByIntReturnsColumn()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Column(1);
            Assert.AreEqual("client 1", column.Data[0]);
            Assert.AreEqual("client 2", column.Data[1]);
        }

        [TestMethod]
        public void ColumnByNameReturnsColumn()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Column("Column 2");
            Assert.AreEqual("client 1", column.Data[0]);
            Assert.AreEqual("client 2", column.Data[1]);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void RowByIntThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Row(10);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void RowByNameThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var row = da.Row("row 5");
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void ColumnByIntThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Column(8);
        }

        [TestMethod]
        [ExpectedException(typeof(IndexOutOfRangeException))]
        public void ColumnByNameThrowsIndexExceptionWhenIndexInvalid()
        {
            DataElement da = Helpers.CreateTestDataElement();
            var column = da.Column("My Column 2");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void RowByNameThrowsArgumentExceptionWhenNoRowHeaders()
        {
            DataElement da = Helpers.CreateTestDataElement();
            da.HasColumnHeaders = false;
            da.HasRowHeaders = false;
            var row = da.Row("row 1");
        }

        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void ColumnByNameThrowsArgumentExceptionWhenNoRowHeaders()
        {
            DataElement da = Helpers.CreateTestDataElement();
            da.HasColumnHeaders = false;
            da.HasRowHeaders = false;
            var column = da.Column("Column 2");
        }

        [TestMethod]
        public void CloneCopiesAllDataToNewObject()
        {
            DataElement da = Helpers.CreateTestDataElement();
            DataElement clone = da.Clone();
            Assert.AreNotSame(da, clone);
            Assert.ReferenceEquals(da, da.Rows[0].ParentElement);
            Assert.ReferenceEquals(da, da.Rows[1].ParentElement);
            Assert.ReferenceEquals(da, da.Columns[0].ParentElement);
            Assert.ReferenceEquals(da, da.Columns[1].ParentElement);
            Assert.ReferenceEquals(da, da.Columns[2].ParentElement);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual(5.05, da.Columns[0].Data[1]);
            Assert.AreEqual("client 1", da.Columns[1].Data[0]);
            Assert.AreEqual("client 2", da.Columns[1].Data[1]);
            Assert.AreEqual(false, da.Columns[2].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[1]);
            Assert.AreEqual(da.Columns[0].Data[0], da.Rows[0].Data[0]);
            Assert.AreEqual(da.Columns[0].Data[1], da.Rows[1].Data[0]);
            Assert.AreEqual(da.Columns[1].Data[0], da.Rows[0].Data[1]);
            Assert.AreEqual(da.Columns[1].Data[1], da.Rows[1].Data[1]);
            Assert.AreEqual(da.Columns[2].Data[0], da.Rows[0].Data[2]);
            Assert.AreEqual(da.Columns[2].Data[1], da.Rows[1].Data[2]);
        }

        [TestMethod]
        public void TrimOrExpandTrimsToDesiredSize()
        {
            DataElement da = Helpers.CreateTestDataElement();
            da.TrimOrExpand(2, 1);
            Assert.AreEqual(2, da.Columns.Count);
            Assert.AreEqual(1, da.Rows.Count);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual("client 1", da.Columns[1].Data[0]);
        }

        [TestMethod]
        public void TrimOrExpandLeavesDataIntactWhenNoSizeChange()
        {
            DataElement da = Helpers.CreateTestDataElement();
            da.TrimOrExpand(3, 2);
            Assert.AreEqual(3, da.Columns.Count);
            Assert.AreEqual(2, da.Rows.Count);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual(5.05, da.Columns[0].Data[1]);
            Assert.AreEqual("client 1", da.Columns[1].Data[0]);
            Assert.AreEqual("client 2", da.Columns[1].Data[1]);
            Assert.AreEqual(false, da.Columns[2].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[1]);
        }

        [TestMethod]
        public void TrimOrExpandAddsEmptyData()
        {
            DataElement da = Helpers.CreateTestDataElement();
            da.TrimOrExpand(4, 3);
            Assert.AreEqual(4, da.Columns.Count);
            Assert.AreEqual(3, da.Rows.Count);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual(5.05, da.Columns[0].Data[1]);
            Assert.IsNull(da.Columns[0].Data[2]);
            Assert.AreEqual("client 1", da.Columns[1].Data[0]);
            Assert.AreEqual("client 2", da.Columns[1].Data[1]);
            Assert.IsNull(da.Columns[1].Data[2]);
            Assert.AreEqual(false, da.Columns[2].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[1]);
            Assert.IsNull(da.Columns[2].Data[2]);
            Assert.IsNull(da.Columns[3].Data[0]);
            Assert.IsNull(da.Columns[3].Data[1]);
            Assert.IsNull(da.Columns[3].Data[2]);
        }

        [TestMethod]
        public void MergeAddsAnotherElement()
        {
            DataElement de1 = Helpers.CreateSmallTestDataElement();
            DataElement de2 = Helpers.CreateSmallTestDataElement();
            de1.Name = "TEST";
            de2.Name = "TEST_B";
            de2.Columns[0].Header = "column A";
            de2.Columns[1].Header = "column B";
            de2.Rows[0].Header = "row A";
            de2.Rows[1].Header = "row B";

            de1.MergeWith(de2);

            Assert.AreEqual(4, de1.Columns.Count);
            Assert.AreEqual(4, de1.Rows.Count);

            Assert.AreEqual(1, de1.Rows[0].Data[0]);
            Assert.AreEqual("A", de1.Rows[0].Data[1]);
            Assert.IsNull(de1.Rows[0].Data[2]);
            Assert.IsNull(de1.Rows[0].Data[3]);

            Assert.AreEqual(2, de1.Rows[1].Data[0]);
            Assert.AreEqual("B", de1.Rows[1].Data[1]);
            Assert.IsNull(de1.Rows[1].Data[2]);
            Assert.IsNull(de1.Rows[1].Data[3]);

            Assert.IsNull(de1.Rows[2].Data[0]);
            Assert.IsNull(de1.Rows[2].Data[1]);
            Assert.AreEqual(1, de1.Rows[2].Data[2]);
            Assert.AreEqual("A", de1.Rows[2].Data[3]);

            Assert.IsNull(de1.Rows[3].Data[0]);
            Assert.IsNull(de1.Rows[3].Data[1]);
            Assert.AreEqual(2, de1.Rows[3].Data[2]);
            Assert.AreEqual("B", de1.Rows[3].Data[3]);
        }

        [TestMethod]
        public void MergeAddsAnotherElementWithOverlapingRowsAndColumns()
        {
            DataElement de1 = Helpers.CreateSmallTestDataElement();
            DataElement de2 = Helpers.CreateSmallTestDataElement();
            de1.Name = "TEST";
            de2.Name = "TEST_B";
            de2.Columns[0].Header = "column A";
            de2.Rows[0].Header = "row A";
            de2.Rows[1].Data[1] = "C";

            de1.MergeWith(de2);

            Assert.AreEqual(3, de1.Columns.Count);
            Assert.AreEqual(3, de1.Rows.Count);

            Assert.AreEqual(1, de1.Rows[0].Data[0]);
            Assert.AreEqual("A", de1.Rows[0].Data[1]);
            Assert.IsNull(de1.Rows[0].Data[2]);

            Assert.AreEqual(2, de1.Rows[1].Data[0]);
            Assert.AreEqual("B", de1.Rows[1].Data[1]);
            Assert.AreEqual(2, de1.Rows[1].Data[2]);

            Assert.IsNull(de1.Rows[2].Data[0]);
            Assert.AreEqual("A", de1.Rows[2].Data[1]);
            Assert.AreEqual(1, de1.Rows[2].Data[2]);
        }

        [TestMethod]
        public void MergeAddsAnotherElementWithOverlapingRowsAndColumnsAndWithoutHeaders()
        {
            DataElement de1 = Helpers.CreateSmallTestDataElement();
            DataElement de2 = Helpers.CreateSmallTestDataElement();
            de1.Name = "TEST";
            de2.Name = "TEST_B";
            de2.Columns[0].Header = null;
            de2.Rows[0].Header = null;
            de2.Rows[1].Data[1] = "C";

            de1.MergeWith(de2);

            Assert.AreEqual(3, de1.Columns.Count);
            Assert.AreEqual(3, de1.Rows.Count);

            Assert.AreEqual(1, de1.Rows[0].Data[0]);
            Assert.AreEqual("A", de1.Rows[0].Data[1]);
            Assert.IsNull(de1.Rows[0].Data[2]);

            Assert.AreEqual(2, de1.Rows[1].Data[0]);
            Assert.AreEqual("B", de1.Rows[1].Data[1]);
            Assert.AreEqual(2, de1.Rows[1].Data[2]);

            Assert.IsNull(de1.Rows[2].Data[0]);
            Assert.AreEqual("A", de1.Rows[2].Data[1]);
            Assert.AreEqual(1, de1.Rows[2].Data[2]);
        }
    }
}
