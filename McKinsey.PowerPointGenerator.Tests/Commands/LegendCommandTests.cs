using System;
using System.Globalization;
using DocumentFormat.OpenXml.Presentation;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;

namespace McKinsey.PowerPointGenerator.Tests.Commands
{
    [TestClass]
    public class LegendCommandTests
    {
        [TestMethod]
        public void ParseArgumentsSetsIndex()
        {
            LegendCommand cmd = new LegendCommand();
            cmd.ArgumentsString = @"""column 1"", ""[value] = 'Q1'"", ""Rectangle 5""";
            var slide = new Slide();
            Shape shape1 = new Shape(@"<p:sp xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""><p:nvSpPr><p:cNvPr id=""20"" name=""Rectangle 4"" /><p:cNvSpPr><a:spLocks noChangeArrowheads=""1"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /></p:cNvSpPr><p:nvPr><p:custDataLst><p:tags r:id=""rId3"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" /></p:custDataLst></p:nvPr></p:nvSpPr><p:spPr bwMode=""gray""><a:xfrm xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:off x=""365919"" y=""1103735"" /><a:ext cx=""8229527"" cy=""4924001"" /></a:xfrm><a:prstGeom prst=""rect"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:avLst /></a:prstGeom><a:solidFill xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""bg1""><a:lumMod val=""95000"" /></a:schemeClr></a:solidFill><a:ln w=""25400"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:noFill /></a:ln><a:effectLst xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:outerShdw blurRad=""50800"" dist=""38100"" dir=""2700000"" algn=""tl"" rotWithShape=""0""><a:prstClr val=""black""><a:alpha val=""40000"" /></a:prstClr></a:outerShdw></a:effectLst></p:spPr><p:style><a:lnRef idx=""2"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""accent1""><a:shade val=""50000"" /></a:schemeClr></a:lnRef><a:fillRef idx=""1"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""accent1"" /></a:fillRef><a:effectRef idx=""0"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""accent1"" /></a:effectRef><a:fontRef idx=""minor"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""lt1"" /></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol=""0"" anchor=""ctr"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /><a:lstStyle xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /><a:p xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:endParaRPr lang=""en-US"" sz=""1100"" b=""1"" dirty=""0""><a:solidFill><a:schemeClr val=""lt1"" /></a:solidFill><a:latin typeface=""+mn-lt"" /></a:endParaRPr></a:p></p:txBody></p:sp>");
            Shape shape2 = new Shape(@"<p:sp xmlns:p=""http://schemas.openxmlformats.org/presentationml/2006/main""><p:nvSpPr><p:cNvPr id=""21"" name=""Rectangle 5"" /><p:cNvSpPr><a:spLocks noChangeArrowheads=""1"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /></p:cNvSpPr><p:nvPr><p:custDataLst><p:tags r:id=""rId4"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"" /></p:custDataLst></p:nvPr></p:nvSpPr><p:spPr bwMode=""gray""><a:xfrm xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:off x=""365919"" y=""1103735"" /><a:ext cx=""8229527"" cy=""4924001"" /></a:xfrm><a:prstGeom prst=""rect"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:avLst /></a:prstGeom><a:solidFill xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""bg1""><a:lumMod val=""95000"" /></a:schemeClr></a:solidFill><a:ln w=""25400"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:noFill /></a:ln><a:effectLst xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:outerShdw blurRad=""50800"" dist=""38100"" dir=""2700000"" algn=""tl"" rotWithShape=""0""><a:prstClr val=""black""><a:alpha val=""40000"" /></a:prstClr></a:outerShdw></a:effectLst></p:spPr><p:style><a:lnRef idx=""2"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""accent1""><a:shade val=""50000"" /></a:schemeClr></a:lnRef><a:fillRef idx=""1"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""accent1"" /></a:fillRef><a:effectRef idx=""0"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""accent1"" /></a:effectRef><a:fontRef idx=""minor"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:schemeClr val=""lt1"" /></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol=""0"" anchor=""ctr"" xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /><a:lstStyle xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main"" /><a:p xmlns:a=""http://schemas.openxmlformats.org/drawingml/2006/main""><a:endParaRPr lang=""en-US"" sz=""1100"" b=""1"" dirty=""0""><a:solidFill><a:schemeClr val=""lt1"" /></a:solidFill><a:latin typeface=""+mn-lt"" /></a:endParaRPr></a:p></p:txBody></p:sp>");
            slide.AppendChild<DocumentFormat.OpenXml.Presentation.Shape>(shape1);
            slide.AppendChild<DocumentFormat.OpenXml.Presentation.Shape>(shape2);
            Document doc = new Document();
            var slideElement = new SlideElement(doc) { Slide = slide };
            ShapeElementBase element = (ShapeElementBase)(ShapeElementBaseTest.Create());
            element.Slide = slideElement;
            cmd.TargetElement = element;

            cmd.ParseArguments();

            Assert.AreEqual("column 1", cmd.Index.Name);
            Assert.AreEqual("[value] = 'Q1'", cmd.Formula);
            Assert.AreEqual("Rectangle 5", cmd.LegendObjectName);
            Assert.AreSame(shape2, cmd.LegendObject.Element);
        }

        [TestMethod]
        public void EvaluateReturnsTrueWhenFormulaIsTrue()
        {
            LegendCommand cmd = new LegendCommand();
            DataElement da = Helpers.CreateTestDataElement();
            cmd.Formula = "[Column 2] == 'client 1'";
            cmd.UsedIndexes.Add(new Index("column 2"));
            var result = FormulaHelper.Evaluate<bool>(da, da.Rows[0], cmd.UsedIndexes, cmd.Formula);
            Assert.IsTrue(result);
        }

        [TestMethod]
        public void EvaluateReturnsFalseWhenFormulaIsFalse()
        {
            LegendCommand cmd = new LegendCommand();
            DataElement da = Helpers.CreateTestDataElement();
            cmd.Formula = "[Column 2] == 'client 12'";
            cmd.UsedIndexes.Add(new Index("column 2"));
            var result = FormulaHelper.Evaluate<bool>(da, da.Rows[0], cmd.UsedIndexes, cmd.Formula);
            Assert.IsFalse(result);
        }

        [TestMethod]
        public void ApplyToDataSetsLegendsValues()
        {
            DataElement da = Helpers.CreateTestDataElement();
            LegendCommand cmd1 = new LegendCommand() { LegendObject = new ShapeElement() { Name = "L1" }, Formula = "[Column 1] > 1", Index = new Index("Column 1") };
            LegendCommand cmd2 = new LegendCommand() { LegendObject = new ShapeElement() { Name = "L2" }, Formula = "[Column 1] <= 1", Index = new Index("Column 1") };
            cmd1.UsedIndexes.Add(cmd1.Index);
            cmd2.UsedIndexes.Add(cmd2.Index);
            cmd1.ApplyToData(da);
            cmd2.ApplyToData(da);
            Assert.AreEqual("L2", ((ShapeElement)da.Columns[0].Legends[0]).Name);
            Assert.AreEqual("L1", ((ShapeElement)da.Columns[0].Legends[1]).Name);
        }
    }
}
