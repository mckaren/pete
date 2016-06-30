using System;
using System.Linq;
using System.Collections.Generic;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Elements
{
    [TestClass]
    public class TableElementTests
    {
        [TestMethod]
        public void ConstructorSeparatesPartsOfTheName()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            TableElement element = TableElement.Create(@"TEST_1[""column 1"", 4, 6][2, 4, ""row 1""] FORMAT{""##,#"", ""de-DE""} FIXED LEGEND_FROM{1} SORT{2, ASC} FORMULA{""'column 2' - 'column 1'""}", null, slide);
            Assert.AreEqual(@"TEST_1[""column 1"", 4, 6][2, 4, ""row 1""] FORMAT{""##,#"", ""de-DE""} FIXED LEGEND_FROM{1} SORT{2, ASC} FORMULA{""'column 2' - 'column 1'""}", element.FullName);
            Assert.AreEqual("TEST_1", element.Name);
            Assert.AreEqual(@"""column 1"", 4, 6", element.DataDescriptor.ColumnIndexesString);
            Assert.AreEqual(@"2, 4, ""row 1""", element.DataDescriptor.RowIndexesString);
            Assert.AreEqual(@"FORMAT{""##,#"", ""de-DE""} FIXED LEGEND_FROM{1} SORT{2, ASC} FORMULA{""'column 2' - 'column 1'""}", element.CommandString);
        }

        [TestMethod]
        public void PreprocessSwitchCommandsGetAllCommandsIfTheyExist()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            TableElement element = TableElement.Create("TEST", null, slide);
            var commands = new List<Command>() { new FormatCommand(), new FormulaCommand(), new NoContentCommand(), new FixedCommand(), new WaterfallCommand(), new PageCommand() };
            var result = element.PreprocessSwitchCommands(commands);
            Assert.AreEqual(5, result.Count());
            Assert.IsTrue(element.IsFixed);
            Assert.IsTrue(element.IsPaged);
        }

        [TestMethod]
        public void PreprocessSwitchCommandsSetsDefaultValuesWhenNoSwitchCommandsExist()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            TableElement element = TableElement.Create("TEST", null, slide);
            var commands = new List<Command>() { new FormatCommand(), new FormulaCommand(), new NoContentCommand()};
            var result = element.PreprocessSwitchCommands(commands);
            Assert.AreEqual(3, result.Count());
            Assert.IsFalse(element.IsFixed);
            Assert.IsFalse(element.IsPaged);
        }
    }
}
