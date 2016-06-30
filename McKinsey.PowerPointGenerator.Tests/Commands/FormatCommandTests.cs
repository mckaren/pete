using System;
using System.Globalization;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class FormatCommandTests
    {
        [TestMethod]
        public void ParseArgumentsSetsFormatAndDefaultCultureWithoutCulture()
        {
            FormatCommand cmd = new FormatCommand();
            cmd.ArgumentsString = @"""##,#""";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual("##,#", cmd.FormatString);
            Assert.AreEqual(CultureInfo.CurrentUICulture.Name, cmd.Culture.Name);
        }

        [TestMethod]
        public void ParseArgumentsSetsFormatAndCultureWithCulture()
        {
            FormatCommand cmd = new FormatCommand();
            cmd.ArgumentsString = @"""##,#"", ""de-DE""";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual("##,#", cmd.FormatString);
            Assert.AreEqual((new CultureInfo("de-DE")).Name, cmd.Culture.Name);
        }

        [TestMethod]
        public void ParseArgumentsSetsFormatAndDefaultCultureWithInvalidCulture()
        {
            FormatCommand cmd = new FormatCommand();
            cmd.ArgumentsString = @"""##,#"", ""aa-aa""";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual("##,#", cmd.FormatString);
            Assert.AreEqual(CultureInfo.CurrentUICulture.Name, cmd.Culture.Name);
        }

        [TestMethod]
        public void ApplyToDataFormatsAllValues()
        {
            DataElement da = Helpers.CreateTestDataElement();
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            element.RowIndexes.Clear();
            element.RowIndexes.Add(new Index(0));
            element.RowIndexes.Add(new Index(1));
            element.ColumnIndexes.Clear();
            element.ColumnIndexes.Add(new Index(0));
            FormatCommand cmd = new FormatCommand();
            cmd.TargetElement = element;
            cmd.FormatString = "0.0000";
            cmd.ApplyToData(da);
            Assert.AreEqual("1.0000", da.Row(0).Data[0]);
            Assert.AreEqual("5.0500", da.Row(1).Data[0]);
            Assert.AreEqual("1.0000", da.Column(0).Data[0]);
            Assert.AreEqual("5.0500", da.Column(0).Data[1]);
        }
    }
}
