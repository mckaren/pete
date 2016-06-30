using System;
using System.Linq;
using System.Collections.Generic;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Elements
{
    [TestClass]
    public class ShapeElementTests
    {
        [TestMethod]
        public void PreprocessSwitchCommandsGetAllCommandsIfTheyExist()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElement element = ShapeElement.Create("TEST", null, slide);
            var commands = new List<Command>() { new FormatCommand(), new FormulaCommand(), new NoContentCommand(), new FixedCommand() };
            var result = element.PreprocessSwitchCommands(commands);
            Assert.AreEqual(3, result.Count());
            Assert.IsTrue(element.IsContentProtected);
        }

        [TestMethod]
        public void PreprocessSwitchCommandsSetsDefaultValuesWhenNoSwitchCommandsExist()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElement element = ShapeElement.Create("TEST", null, slide);
            var commands = new List<Command>() { new FormatCommand(), new FormulaCommand(), new FixedCommand() };
            var result = element.PreprocessSwitchCommands(commands);
            Assert.AreEqual(3, result.Count());
            Assert.IsFalse(element.IsContentProtected);
        }
    }
}
