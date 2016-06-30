using System;
using System.Linq;
using System.Collections.Generic;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using McKinsey.PowerPointGenerator.Core.Data;
using Moq;

namespace McKinsey.PowerPointGenerator.Tests.Elements
{
    [TestClass]
    public class ShepeElementBaseTests
    {
        [TestMethod]
        public void ParseSeparatesPartsOfTheNameStringAndExtractsIndexes()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create(@"TEST_1[""column 1"", 4, 6][2, 4, ""row 1""] FORMAT{""##,#"", ""de-DE""} FIXED LEGEND_FROM{1} SORT{2, ASC} FORMULA{""[column 2] - [column 1]""}", slide));
            Assert.AreEqual(@"TEST_1[""column 1"", 4, 6][2, 4, ""row 1""] FORMAT{""##,#"", ""de-DE""} FIXED LEGEND_FROM{1} SORT{2, ASC} FORMULA{""[column 2] - [column 1]""}", element.FullName);
            Assert.AreEqual("TEST_1", element.Name);
            Assert.AreEqual(@"""column 1"", 4, 6", element.DataDescriptor.ColumnIndexesString);
            Assert.AreEqual(@"2, 4, ""row 1""", element.DataDescriptor.RowIndexesString);
            Assert.AreEqual(@"FORMAT{""##,#"", ""de-DE""} FIXED LEGEND_FROM{1} SORT{2, ASC} FORMULA{""[column 2] - [column 1]""}", element.CommandString);
            Assert.AreEqual("column 1", element.ColumnIndexes[0].Name);
            Assert.AreEqual(4, element.ColumnIndexes[1].Number.Value);
            Assert.AreEqual(6, element.ColumnIndexes[2].Number.Value);
            Assert.AreEqual(2, element.RowIndexes[0].Number.Value);
            Assert.AreEqual(4, element.RowIndexes[1].Number.Value);
            Assert.AreEqual("row 1", element.RowIndexes[2].Name);
        }

        [TestMethod]
        public void ParseRecognisesAdditionalDataElements()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create(@"TEST_1[""column 1"", 4, 6][2, 4, ""row 1""] TEST_A[""column A1""][""row A2""] TEST_B[""column B1""][""row B2""] NO_CONTENT FORMAT{""##,#"", ""de-DE""} FIXED LEGEND_FROM{1} SORT{2, ASC} FORMULA{""[column 2] - [column 1]""}", slide));
            Assert.AreEqual(@"TEST_1[""column 1"", 4, 6][2, 4, ""row 1""] TEST_A[""column A1""][""row A2""] TEST_B[""column B1""][""row B2""] NO_CONTENT FORMAT{""##,#"", ""de-DE""} FIXED LEGEND_FROM{1} SORT{2, ASC} FORMULA{""[column 2] - [column 1]""}", element.FullName);
            Assert.AreEqual("TEST_1", element.Name);
            Assert.AreEqual(@"""column 1"", 4, 6", element.DataDescriptor.ColumnIndexesString);
            Assert.AreEqual(@"2, 4, ""row 1""", element.DataDescriptor.RowIndexesString);
            Assert.AreEqual(2, element.AdditionalDataDescriptors.Count);
            Assert.AreEqual(@"""column A1""", element.AdditionalDataDescriptors[0].ColumnIndexesString);
            Assert.AreEqual(@"""row A2""", element.AdditionalDataDescriptors[0].RowIndexesString);
            Assert.AreEqual("TEST_A", element.AdditionalDataDescriptors[0].Name);
            Assert.AreEqual(@"""column B1""", element.AdditionalDataDescriptors[1].ColumnIndexesString);
            Assert.AreEqual(@"""row B2""", element.AdditionalDataDescriptors[1].RowIndexesString);
            Assert.AreEqual("TEST_B", element.AdditionalDataDescriptors[1].Name);
            Assert.AreEqual(@"NO_CONTENT FORMAT{""##,#"", ""de-DE""} FIXED LEGEND_FROM{1} SORT{2, ASC} FORMULA{""[column 2] - [column 1]""}", element.CommandString);
            Assert.AreEqual("column 1", element.ColumnIndexes[0].Name);
            Assert.AreEqual(4, element.ColumnIndexes[1].Number.Value);
            Assert.AreEqual(6, element.ColumnIndexes[2].Number.Value);
            Assert.AreEqual(@"column A1", element.ColumnIndexes[3].Name);
            Assert.AreEqual(@"column B1", element.ColumnIndexes[4].Name);
            Assert.AreEqual(2, element.RowIndexes[0].Number.Value);
            Assert.AreEqual(4, element.RowIndexes[1].Number.Value);
            Assert.AreEqual("row 1", element.RowIndexes[2].Name);
            Assert.AreEqual("row A2", element.RowIndexes[3].Name);
            Assert.AreEqual("row B2", element.RowIndexes[4].Name);
        }

        [TestMethod]
        public void ParseBehavesWellWhenThereAreNoCommandsOrIndexes()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create(@"TEST_1", slide));
            Assert.AreEqual(@"TEST_1", element.FullName);
            Assert.AreEqual("TEST_1", element.Name);
            Assert.AreEqual("", element.DataDescriptor.ColumnIndexesString);
            Assert.AreEqual("", element.DataDescriptor.RowIndexesString);
            Assert.AreEqual("", element.CommandString);
        }

        [TestMethod]
        public void PreprocessSwitchCommandsBaseSetsRowAndColumnHeadersAndRemovesCommandsFromTheList()
        {
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            var commands = new List<Command>() { new FormatCommand(), new RowHeaderCommand(), new FormulaCommand(), new FixedCommand(), new ColumnHeaderCommand() };
            var result = element.PreprocessSwitchCommandsBase(commands);
            Assert.AreEqual(3, result.Count());
            Assert.IsTrue(element.UseRowHeaders);
            Assert.IsTrue(element.UseColumnHeaders);
        }

        [TestMethod]
        public void PreprocessSwitchCommandsBaseSetsRowAndColumnHeadersToTrueWhenNoCommandsAreOnTheList()
        {
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            var commands = new List<Command>() { new FormatCommand(), new FormulaCommand(), new FixedCommand() };
            var result = element.PreprocessSwitchCommandsBase(commands);
            Assert.AreEqual(3, result.Count());
            Assert.IsFalse(element.UseRowHeaders);
            Assert.IsFalse(element.UseColumnHeaders);
        }

        [TestMethod]
        public void FindShapeDataSetsDataCloneWhenMatchingElementFoundAndReturnsTrue()
        {
            DataElement de1 = new DataElement { Name = "test 2" };
            DataElement de2 = new DataElement { Name = "Test" };
            IList<DataElement> data = new List<DataElement> { de1, de2 };
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create("TEST"));

            element.FindShapeData(data);
            Assert.AreNotSame(de2, element.Data);
        }

        [TestMethod]
        public void FindShapeSetsDataToNullWhenMatchingElementNotFound()
        {
            DataElement de1 = new DataElement { Name = "test 2" };
            DataElement de2 = new DataElement { Name = "Test 1" };
            IList<DataElement> data = new List<DataElement> { de1, de2 };
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create("TEST"));
            element.Slide = new SlideElement(new Document()) { Number = 1 };

            element.FindShapeData(data);
            Assert.IsNull(element.Data);
        }

        [TestMethod]
        public void FindShapeCallsMergeWithOnAdditionalDataElements()
        {
            var de1 = new Mock<DataElement>();
            var de2 = new Mock<DataElement>();
            de1.SetupGet(e => e.Name).Returns("TEST");
            de1.Setup(e => e.Clone()).Returns(de1.Object);
            de2.SetupGet(e => e.Name).Returns("TEST_B");
            de2.Setup(e => e.Clone()).Returns(de2.Object);
            IList<DataElement> data = new List<DataElement> { de1.Object, de2.Object };
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create("TEST"));
            element.AdditionalDataDescriptors.Add(new DataElementDescriptor { Name = "TEST_B" });
            element.FindShapeData(data);

            de1.Verify(e => e.MergeWith(It.IsAny<DataElement>()), Times.Once());
            de2.Verify(e => e.MergeWith(It.IsAny<DataElement>()), Times.Never());
        }

        [TestMethod]
        public void CheckCommandsForIndexesAddsIndexesToTheList()
        {
            var cmd1 = new FormulaCommand();
            cmd1.UsedIndexes.Add(new Index("#Cl1N#"));
            cmd1.UsedIndexes.Add(new Index("#Cl2N#"));
            var cmd2 = new FormulaCommand();
            cmd2.UsedIndexes.Add(new Index("#P1N#"));
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create("TEST"));
            element.Commands.Add(cmd1);
            element.Commands.Add(cmd2);
            element.CheckCommandsForIndexes();
            Assert.AreEqual("#Cl1N#", element.ColumnIndexes[0].Name);
            Assert.AreEqual("#Cl2N#", element.ColumnIndexes[1].Name);
            Assert.AreEqual("#P1N#", element.ColumnIndexes[2].Name);
            Assert.IsTrue(element.ColumnIndexes[0].IsHidden);
            Assert.IsTrue(element.ColumnIndexes[1].IsHidden);
            Assert.IsTrue(element.ColumnIndexes[2].IsHidden);
        }

        [TestMethod]
        public void CheckCommandsForIndexesAddsIndexesToTheListButLeavesExistingAlone()
        {
            var cmd1 = new FormulaCommand();
            cmd1.UsedIndexes.Add(new Index("#Cl1N#"));
            cmd1.UsedIndexes.Add(new Index("#Cl2N#"));
            var cmd2 = new FormulaCommand();
            cmd2.UsedIndexes.Add(new Index("#P1N#"));
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create("TEST"));
            element.ColumnIndexes.Add(new Index("#Cl1N#"));
            element.Commands.Add(cmd1);
            element.Commands.Add(cmd2);
            element.CheckCommandsForIndexes();
            Assert.AreEqual("#Cl1N#", element.ColumnIndexes[0].Name);
            Assert.AreEqual("#Cl2N#", element.ColumnIndexes[1].Name);
            Assert.AreEqual("#P1N#", element.ColumnIndexes[2].Name);
            Assert.IsFalse(element.ColumnIndexes[0].IsHidden);
            Assert.IsTrue(element.ColumnIndexes[1].IsHidden);
            Assert.IsTrue(element.ColumnIndexes[2].IsHidden);
        }

        [TestMethod]
        public void CheckCommandsForIndexesAddsIndexesToTheListButUsesRowsIfTransposed()
        {
            var sc = new FormulaCommand();
            sc.UsedIndexes.Add(new Index("Totals"));
            var tc = new TransposeCommand();
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create("TEST"));
            element.ColumnIndexes.Add(new Index("#Cl1N#"));
            element.ColumnIndexes.Add(new Index("#Cl2N#"));
            element.RowIndexes.Add(new Index("AD"));
            element.RowIndexes.Add(new Index("AM"));
            element.Commands.Add(tc);
            element.Commands.Add(sc);
            element.CheckCommandsForIndexes();
            Assert.AreEqual(2, element.ColumnIndexes.Count);
            Assert.AreEqual(3, element.RowIndexes.Count);
            Assert.AreEqual("#Cl1N#", element.ColumnIndexes[0].Name);
            Assert.AreEqual("#Cl2N#", element.ColumnIndexes[1].Name);
            Assert.AreEqual("AD", element.RowIndexes[0].Name);
            Assert.AreEqual("AM", element.RowIndexes[1].Name);
            Assert.AreEqual("Totals", element.RowIndexes[2].Name);
            Assert.IsFalse(element.ColumnIndexes[0].IsHidden);
            Assert.IsFalse(element.ColumnIndexes[1].IsHidden);
            Assert.IsFalse(element.RowIndexes[0].IsHidden);
            Assert.IsFalse(element.RowIndexes[1].IsHidden);
            Assert.IsTrue(element.RowIndexes[2].IsHidden);
        }

        [TestMethod]
        public void ParseExtractsIndexesWithRanges()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create(@"TEST_1[2, ""column 1"", 4-6][2-4, 8, ""row 1""]", slide));

            Assert.AreEqual(2, element.ColumnIndexes[0].Number.Value);
            Assert.AreEqual("column 1", element.ColumnIndexes[1].Name);
            Assert.AreEqual(4, element.ColumnIndexes[2].Number.Value);
            Assert.AreEqual(5, element.ColumnIndexes[3].Number.Value);
            Assert.AreEqual(6, element.ColumnIndexes[4].Number.Value);

            Assert.AreEqual(2, element.RowIndexes[0].Number.Value);
            Assert.AreEqual(3, element.RowIndexes[1].Number.Value);
            Assert.AreEqual(4, element.RowIndexes[2].Number.Value);
            Assert.AreEqual(8, element.RowIndexes[3].Number.Value);
            Assert.AreEqual("row 1", element.RowIndexes[4].Name);
        }
    }
}
