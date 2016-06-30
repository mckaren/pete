using System;
using System.Globalization;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class ReplaceCommandTests
    {
        [TestMethod]
        public void ParseArgumentsSetsReplacements()
        {
            ReplaceCommand cmd = new ReplaceCommand();
            cmd.ArgumentsString = @"""true"" = ""ü"", ""false""="""", ""1""=""one"", ""client""=""IBM""";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual(4, cmd.Substitutions.Count);
            Assert.AreEqual("ü", cmd.Substitutions["true"]);
            Assert.AreEqual("", cmd.Substitutions["false"]);
            Assert.AreEqual("one", cmd.Substitutions["1"]);
            Assert.AreEqual("IBM", cmd.Substitutions["client"]);
        }

        [TestMethod]
        public void ApplyToDataReplacesValues()
        {
            ReplaceCommand cmd = new ReplaceCommand();
            cmd.Substitutions.Add("client 1", "AMD");
            cmd.Substitutions.Add("client 2", "Intel");
            DataElement data = Helpers.CreateTestDataElementWithTwoNumericColumns();
            cmd.ApplyToData(data);
            Assert.AreEqual("AMD", data.Columns[1].Data[0]);
            Assert.AreEqual("Intel", data.Columns[1].Data[1]);
        }
    }
}
