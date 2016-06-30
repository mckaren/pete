using System;
using System.Linq;
using System.Collections.Generic;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Elements
{
    [TestClass]
    public class TextElementTests
    {
        [TestMethod]
        public void ConstructorAddsFormatCommandIfFormatIsProvided()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            TextElement element = TextElement.Create(@"Some_number[1][4]:{""##,#"", ""en-US""}", null, slide);
            Assert.AreEqual(@"Some_number[1][4]:{""##,#"", ""en-US""}", element.FullName);
            Assert.AreEqual("Some_number", element.Name);
            Assert.AreEqual(@"1", element.DataDescriptor.ColumnIndexesString);
            Assert.AreEqual(@"4", element.DataDescriptor.RowIndexesString);
            Assert.AreEqual(@"FORMAT{""##,#"", ""en-US""}", element.CommandString);
        }

        [TestMethod]
        public void ConstructorAddsGetsTextIndexes()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            TextElement element = TextElement.Create(@"Some_number[""column 1""][""row 1""]:{""##,#"", ""en-US""}", null, slide);
            Assert.AreEqual(@"Some_number[""column 1""][""row 1""]:{""##,#"", ""en-US""}", element.FullName);
            Assert.AreEqual("Some_number", element.Name);
            Assert.AreEqual(@"""column 1""", element.DataDescriptor.ColumnIndexesString);
            Assert.AreEqual(@"""row 1""", element.DataDescriptor.RowIndexesString);
            Assert.AreEqual(@"FORMAT{""##,#"", ""en-US""}", element.CommandString);
        }

        [TestMethod]
        public void PreprocessSwitchCommandsNotChangingCommands()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            TextElement element = TextElement.Create("TEST", null, slide);
            var commands = new List<Command>() { new FormatCommand(), new FormulaCommand(), new NoContentCommand() };
            var result = element.PreprocessSwitchCommands(commands);
            Assert.AreEqual(3, result.Count());
        }
    }
}
