using System;
using System.Globalization;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class VisibleCommandTests
    {
        [TestMethod]
        public void ParseArgumentsSetsFormula()
        {
            VisibleCommand cmd = new VisibleCommand();
            cmd.ArgumentsString = @"""[value] > 5 OR [value] <- 5""";
            Document doc = new Document();
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual("[value] > 5 OR [value] <- 5", cmd.Formula);
        }

        [TestMethod]
        public void ParseArgumentsSetsUsedIndexes()
        {
            VisibleCommand cmd = new VisibleCommand();
            cmd.ArgumentsString = @"""[column 1] > 5 OR [3] <- 5""";
            Document doc = new Document();
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual(2, cmd.UsedIndexes.Count);
            Assert.AreEqual("column 1", cmd.UsedIndexes[0].Name);
            Assert.AreEqual(3, cmd.UsedIndexes[1].Number.Value);
        }

        [TestMethod]
        public void ApplyToDataSetsVisibleToTrueWhenFormulaReturnsTrue()
        {
            VisibleCommand cmd = new VisibleCommand();
            cmd.Formula = @"[column 1] > 5 OR [c#1] <= 5";
            cmd.UsedIndexes.Add(new Index("column 1"));
            cmd.UsedIndexes.Add(new Index(1));
            var data = Helpers.CreateTwoValueElement("data", 1, 5);
            data.Columns[0].Header = "column 1";
            data.HasColumnHeaders = true;
            cmd.ApplyToData(data);
            Assert.IsTrue(cmd.IsVisible);
        }

        [TestMethod]
        public void ApplyToDataSetsVisibleToFalseWhenFormulaReturnsFalse()
        {
            VisibleCommand cmd = new VisibleCommand();
            cmd.Formula = @"[column 1] > 5 OR [c#1] <= -5";
            cmd.UsedIndexes.Add(new Index("column 1"));
            cmd.UsedIndexes.Add(new Index(1));
            var data = Helpers.CreateTwoValueElement("data", 1, 5);
            data.Columns[0].Header = "column 1";
            data.HasColumnHeaders = true;
            cmd.ApplyToData(data);
            Assert.IsFalse(cmd.IsVisible);
        }


        [TestMethod]
        public void ApplyToDataSetsVisibleToFalseWhenFormulaReturnsFalseWithSingleValueElement()
        {
            VisibleCommand cmd = new VisibleCommand();
            cmd.Formula = @"[value] > 1";
            cmd.UsedIndexes.Add(new Index("column 1"));
            cmd.UsedIndexes.Add(new Index(1));
            var data = Helpers.CreateSingleValueElement("data", 1);
            data.Columns[0].Header = "column 1";
            data.HasColumnHeaders = true;
            cmd.ApplyToData(data);
            Assert.IsFalse(cmd.IsVisible);
        }


        [TestMethod]
        public void ApplyToDataSetsVisibleToTrueWhenFormulaReturnsTrueWithSingleValueElement()
        {
            VisibleCommand cmd = new VisibleCommand();
            cmd.Formula = @"[value] >= 1";
            var data = Helpers.CreateSingleValueElement("data", 1);
            data.Columns[0].Header = "column 1";
            data.HasColumnHeaders = true;
            cmd.ApplyToData(data);
            Assert.IsTrue(cmd.IsVisible);
        }
    }
}
