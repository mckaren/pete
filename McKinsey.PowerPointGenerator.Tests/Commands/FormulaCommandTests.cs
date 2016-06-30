using System;
using System.Globalization;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class FormulaCommandTests
    {
        [TestMethod]
        public void ParseArgumentsSetsFormula()
        {
            FormulaCommand cmd = new FormulaCommand();
            cmd.ArgumentsString = @"""[column 2] - [column 1]""";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual("[column 2] - [column 1]", cmd.Formula);
            Assert.AreEqual("column 2", cmd.UsedIndexes[0].Name);
            Assert.AreEqual("column 1", cmd.UsedIndexes[1].Name);
            Assert.IsFalse(cmd.IsInPlaceFormula);
        }

        [TestMethod]
        public void ParseArgumentsSetsFormulaWithStringValues()
        {
            FormulaCommand cmd = new FormulaCommand();
            cmd.ArgumentsString = @"""[column 2] = 'test'""";
            Document doc = new Document();
            var slide = new SlideElement(doc);
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            cmd.TargetElement = element;
            cmd.ParseArguments();
            Assert.AreEqual("[column 2] = 'test'", cmd.Formula);
            Assert.AreEqual("column 2", cmd.UsedIndexes[0].Name);
            Assert.AreEqual(1, cmd.UsedIndexes.Count);
            Assert.IsFalse(cmd.IsInPlaceFormula);
        }

        [TestMethod]
        public void FormulaCommandEvaluatesTheExpression()
        {
            FormulaCommand cmd = new FormulaCommand() { Formula = @"[column 4] - [column 1]" };
            DataElement data = Helpers.CreateTestDataElementWithTwoNumericColumns();
            cmd.UsedIndexes.Add(new Index("column 4"));
            cmd.UsedIndexes.Add(new Index("column 1"));
            cmd.ApplyToData(data);
            Assert.AreEqual("Calc1", data.Columns[4].Header);
            Assert.AreEqual(1.5, data.Columns[4].Data[0]);
            Assert.AreEqual(-1.0, data.Columns[4].Data[1]);
        }


        [TestMethod]
        public void FormulaCommandEvaluatesTheExpressionWithString()
        {
            FormulaCommand cmd = new FormulaCommand() { Formula = @"[Column 2] == 'client 1'" };
            DataElement data = Helpers.CreateTestDataElementWithTwoNumericColumns();
            cmd.UsedIndexes.Add(new Index("column 2"));
            cmd.ApplyToData(data);
            Assert.AreEqual("Calc1", data.Columns[4].Header);
            Assert.AreEqual(true, data.Columns[4].Data[0]);
            Assert.AreEqual(false, data.Columns[4].Data[1]);
        }

        [TestMethod]
        public void MultipleFormulaCommandsEvaluatesOnPreviousResult()
        {
            FormulaCommand cmd = new FormulaCommand() { Formula = @"[column 4] - [column 1]" };
            DataElement data = Helpers.CreateTestDataElementWithTwoNumericColumns();
            cmd.UsedIndexes.Add(new Index("column 4"));
            cmd.UsedIndexes.Add(new Index("column 1"));
            cmd.ApplyToData(data);
            FormulaCommand cmd1 = new FormulaCommand() { Formula = @"[calc1] * -2" };
            cmd1.UsedIndexes.Add(new Index("calc1"));
            cmd1.ApplyToData(data);
            Assert.AreEqual("Calc1", data.Columns[4].Header);
            Assert.AreEqual(1.5, data.Columns[4].Data[0]);
            Assert.AreEqual(-1.0, data.Columns[4].Data[1]);
            Assert.AreEqual("Calc2", data.Columns[5].Header);
            Assert.AreEqual(-3.0, data.Columns[5].Data[0]);
            Assert.AreEqual(2.0, data.Columns[5].Data[1]);
        }

        [TestMethod]
        public void FormulaCommandEvaluatesTheExpressionWithValue()
        {
            FormulaCommand cmd = new FormulaCommand() { Formula = @"[value] * 1000" };
            DataElement data = Helpers.CreateSingleValueElement("test", 5);
            cmd.IsInPlaceFormula = true;
            cmd.ApplyToData(data);
            Assert.AreEqual(5000, data.Columns[0].Data[0]);
        }
    }
}
