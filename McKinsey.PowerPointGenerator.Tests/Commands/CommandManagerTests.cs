using System;
using System.Linq;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using McKinsey.PowerPointGenerator.Commands;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class CommandManagerTests
    {
        [TestMethod]
        public void DiscoverCommandsResolvesCommandsAndSetsProcessingOrder()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElement element = ShapeElement.Create("TEST FORMAT(\"##,#\", \"de-DE\") FIXED TRANSPOSE LEGEND(\"#Cl1N#\", \"G1=='Q1'\", \"Q1\") SORT(2, ASC) FORMULA(\"'column 2' - 'column 1'\")", null, slide);
            var commands = CommandManager.DiscoverCommands(element).ToList();
            Assert.AreEqual(6, commands.Count);
            Assert.IsInstanceOfType(commands[0], typeof(FormatCommand));
            Assert.IsInstanceOfType(commands[1], typeof(FixedCommand));
            Assert.IsInstanceOfType(commands[2], typeof(TransposeCommand));
            Assert.IsInstanceOfType(commands[3], typeof(LegendCommand));
            Assert.IsInstanceOfType(commands[4], typeof(SortCommand));
            Assert.IsInstanceOfType(commands[5], typeof(FormulaCommand));
        }

        [TestMethod]
        public void DiscoverCommandsSetsArgumentsString()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElement element = ShapeElement.Create("TEST FORMAT(\"##,#\", \"de-DE\") FIXED TRANSPOSE LEGEND(\"#Cl1N#\", \"G1=='Q1'\", \"Q1\") SORT(2, ASC) FORMULA(\"'column 2' - 'column 1'\")", null, slide);
            var commands = CommandManager.DiscoverCommands(element).ToList();
            Assert.AreEqual(6, commands.Count);
            Assert.AreEqual("\"##,#\", \"de-DE\"", commands[0].ArgumentsString);
            Assert.AreEqual("", commands[1].ArgumentsString);
            Assert.AreEqual("", commands[2].ArgumentsString);
            Assert.AreEqual("\"#Cl1N#\", \"G1=='Q1'\", \"Q1\"", commands[3].ArgumentsString);
            Assert.AreEqual("2, ASC", commands[4].ArgumentsString);
            Assert.AreEqual("\"'column 2' - 'column 1'\"", commands[5].ArgumentsString);
        }
       
        [TestMethod]
        public void DiscoverCommandsReturnsEmptyListWhenNoCommandsAreFound()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElement element = ShapeElement.Create("TEST", null, slide);
            var commands = CommandManager.DiscoverCommands(element).ToList();
            Assert.AreEqual(0, commands.Count);
        }

        [TestMethod]
        public void DiscoverCommandsIgnorwsUnknownCommands()
        {
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElement element = ShapeElement.Create("TEST ADD_COLUMN_HEADER", null, slide);
            var commands = CommandManager.DiscoverCommands(element).ToList();
            Assert.AreEqual(0, commands.Count);
        }
    }
}
