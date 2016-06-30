using System;
using System.Collections.Generic;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Processing;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Processing
{
    [TestClass]
    public class DataElementProcessorTests
    {
        [TestMethod]
        public void ParseAndReplaceReplacesAllOccurencesOfTag()
        {
            var clientName = Helpers.CreateSingleValueElement("client_name", "Intel");
            var competitorName = Helpers.CreateSingleValueElement("competitor_name", "AMD");
            var header = Helpers.CreateSingleValueElement("header", "#client_name# vs. #competitor_name# performance in 2014");
            List<DataElement> dataSet = new List<DataElement> { clientName, header, competitorName };
            var result = DataElementProcessor.ParseAndReplace((string)header.Rows[0].Data[0], dataSet);
            Assert.AreEqual("Intel vs. AMD performance in 2014", result);
        }

        [TestMethod]
        public void ParseAndReplaceKeepsTagWhenElementNotFound()
        {
            var clientName = Helpers.CreateSingleValueElement("client_name", "Intel");
            var competitorName = Helpers.CreateSingleValueElement("competitor_name", "AMD");
            var header = Helpers.CreateSingleValueElement("header", "#client_name# vs. #main_competitor_name# performance in 2014");
            List<DataElement> dataSet = new List<DataElement> { clientName, header, competitorName };
            var result = DataElementProcessor.ParseAndReplace((string)header.Rows[0].Data[0], dataSet);
            Assert.AreEqual("Intel vs. #main_competitor_name# performance in 2014", result);
        }

        [TestMethod]
        public void ParseAndReplaceReplacesAllOccurencesOfTagWithIndexes()
        {
            var clientName = Helpers.CreateSingleValueElement("client_name", "Intel");
            var competitorName = Helpers.CreateSingleValueElement("competitor_name", "AMD");
            var benchmark = Helpers.CreateTwoValueElement("benchmark", "GPU Benchmark", 2014);
            var header = Helpers.CreateSingleValueElement("header", "#client_name# vs. #competitor_name# performance in #benchmark[1][0]#");
            List<DataElement> dataSet = new List<DataElement> { clientName, header, competitorName, benchmark };
            var result = DataElementProcessor.ParseAndReplace((string)header.Rows[0].Data[0], dataSet);
            Assert.AreEqual("Intel vs. AMD performance in 2014", result);
        }

        [TestMethod]
        public void ParseAndReplaceReplacesAllOccurencesOfTagWithFormat()
        {
            var clientName = Helpers.CreateSingleValueElement("client_name", "Intel");
            var competitorName = Helpers.CreateSingleValueElement("competitor_name", "AMD");
            var benchmark = Helpers.CreateSingleValueElement("benchmark", new DateTime(2013, 01, 01));
            var header = Helpers.CreateSingleValueElement("header", "#client_name# vs. #competitor_name# performance as of #benchmark:{dd MMM yyyy}#");
            List<DataElement> dataSet = new List<DataElement> { clientName, header, competitorName, benchmark };
            var result = DataElementProcessor.ParseAndReplace((string)header.Rows[0].Data[0], dataSet);
            Assert.AreEqual("Intel vs. AMD performance as of 01 Jan 2013", result);
        }

        [TestMethod]
        public void ProcessReplacesElements()
        {
            var clientName = Helpers.CreateSingleValueElement("client_name", "Intel");
            var competitorName = Helpers.CreateSingleValueElement("competitor_name", "AMD");
            var header = Helpers.CreateSingleValueElement("header", "#client_name# vs. #competitor_name# performance as of #benchmark:{dd MMM yyyy}#");
            var benchmark = Helpers.CreateSingleValueElement("benchmark", new DateTime(2013, 01, 01));
            var da = Helpers.CreateTestDataElement();
            da.Rows[0].Data[1] = "#client_name#";
            da.Rows[1].Data[1] = "#competitor_name#";
            da.Rows[0].Header = "#client_name#";
            da.Rows[1].Header = "#competitor_name#";
            da.Columns[1].Header = "#benchmark:{d MMMM yyyy}#";
            da.Columns[0].Data.Clear();
            da.Columns[1].Data.Clear();
            da.Columns[2].Data.Clear();
            da.Normalize();
            List<DataElement> dataSet = new List<DataElement> { clientName, da, competitorName, header, benchmark };
            DataElementProcessor.Process(dataSet);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual(5.05, da.Columns[0].Data[1]);
            Assert.AreEqual("Intel", da.Columns[1].Data[0]);
            Assert.AreEqual("AMD", da.Columns[1].Data[1]);
            Assert.AreEqual(false, da.Columns[2].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[1]);
            Assert.AreEqual("Intel", da.Rows[0].MappedHeader);
            Assert.AreEqual("#client_name#", da.Rows[0].Header);
            Assert.AreEqual("AMD", da.Rows[1].MappedHeader);
            Assert.AreEqual("#competitor_name#", da.Rows[1].Header);
            Assert.AreEqual("1 January 2013", da.Columns[1].MappedHeader);
            Assert.AreEqual("#benchmark:{d MMMM yyyy}#", da.Columns[1].Header);
            Assert.AreEqual("Intel vs. AMD performance as of 01 Jan 2013", header.Rows[0].Data[0]);
        }

        [TestMethod]
        public void ProcessSkipsWhenReplacementsAreNull()
        {
            var clientName = Helpers.CreateSingleValueElement("client_name", null);
            var competitorName = Helpers.CreateSingleValueElement("competitor_name", "AMD");
            var header = Helpers.CreateSingleValueElement("header", "#client_name# vs. #competitor_name# performance as of #benchmark:{dd MMM yyyy}#");
            var benchmark = Helpers.CreateSingleValueElement("benchmark", new DateTime(2013, 01, 01));
            var da = Helpers.CreateTestDataElement();
            da.Rows[0].Data[1] = "#client_name#";
            da.Rows[1].Data[1] = "#competitor_name#";
            da.Rows[0].Header = "#client_name#";
            da.Rows[1].Header = "#competitor_name#";
            da.Columns[1].Header = "#benchmark:{d MMMM yyyy}#";
            da.Columns[0].Data.Clear();
            da.Columns[1].Data.Clear();
            da.Columns[2].Data.Clear();
            da.Normalize();
            List<DataElement> dataSet = new List<DataElement> { clientName, da, competitorName, header, benchmark };
            DataElementProcessor.Process(dataSet);
            Assert.AreEqual(1.0, da.Columns[0].Data[0]);
            Assert.AreEqual(5.05, da.Columns[0].Data[1]);
            Assert.AreEqual("#client_name#", da.Columns[1].Data[0]);
            Assert.AreEqual("AMD", da.Columns[1].Data[1]);
            Assert.AreEqual(false, da.Columns[2].Data[0]);
            Assert.AreEqual(true, da.Columns[2].Data[1]);
            Assert.AreEqual("#client_name#", da.Rows[0].Header);
            Assert.AreEqual("Totals", da.Rows[0].MappedHeader);
            Assert.AreEqual("#competitor_name#", da.Rows[1].Header);
            Assert.AreEqual("AMD", da.Rows[1].MappedHeader);
            Assert.AreEqual("#benchmark:{d MMMM yyyy}#", da.Columns[1].Header);
            Assert.AreEqual("1 January 2013", da.Columns[1].MappedHeader);
            Assert.AreEqual("#client_name# vs. AMD performance as of 01 Jan 2013", header.Rows[0].Data[0]);
        }

        [TestMethod]
        public void ProcessDataReplacesDeepReferences()
        {
            var clientName = Helpers.CreateSingleValueElement("client_name", "IBM #country#");
            var country = Helpers.CreateSingleValueElement("country", "UK");
            var benchmark = Helpers.CreateSingleValueElement("benchmark", "benchmark for #client_name#");
            List<DataElement> dataSet = new List<DataElement> { clientName, country, benchmark };
            DataElementProcessor.Process(dataSet);
            Assert.AreEqual("IBM UK", clientName.Rows[0].Data[0]);
            Assert.AreEqual("UK", country.Rows[0].Data[0]);
            Assert.AreEqual("benchmark for IBM UK", benchmark.Rows[0].Data[0]);
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void ProcessDataPreventsCircularReferences()
        {
            var clientName = Helpers.CreateSingleValueElement("client_name", "this is #competitor_name#");
            var competitorName = Helpers.CreateSingleValueElement("competitor_name", "this is #client_name#");
            var benchmark = Helpers.CreateSingleValueElement("benchmark", new DateTime(2013, 01, 01));
            List<DataElement> dataSet = new List<DataElement> { clientName, competitorName, benchmark };
            DataElementProcessor.Process(dataSet);
        }

        [TestMethod]
        public void GetFragmentByIndexesIgnoresNonExistingColumns()
        {
            var da = Helpers.CreateTestDataElement();
            var daf = da.GetFragmentByIndexes(new List<Index> { new Index("row 1") }, new List<Index> { new Index("column 5") });
            Assert.AreEqual(1, daf.Rows.Count);
            Assert.AreEqual(3, daf.Columns.Count);
        }

        [TestMethod]
        public void GetFragmentByIndexesIgnoresNonExistingRows()
        {
            var da = Helpers.CreateTestDataElement();
            var daf = da.GetFragmentByIndexes(new List<Index> { new Index("row 5") }, new List<Index> { new Index("column 1") });
            Assert.AreEqual(2, daf.Rows.Count);
            Assert.AreEqual(1, daf.Columns.Count);
        }
    }
}
