using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using McKinsey.PowerPointGenerator.Extensions;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing.Charts;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Core.Data;
using NLog;

namespace McKinsey.PowerPointGenerator
{
    [DebuggerDisplay("slide, number: {Number}")]
    public class SlideElement
    {
        private static Regex textTagRegex = new Regex(@"#(?<tag>.*?(?:\[.*?\])?(?:\[.*?\])?(?:\:\(.*?\))?)#", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public int Number { get; set; }
        public virtual Slide Slide { get; set; }
        public virtual List<ShapeElementBase> Shapes { get; internal set; }
        public Document Document { get; set; }

        public SlideElement(Document document)
        {
            Document = document;
            Shapes = new List<ShapeElementBase>();
        }

        public virtual void DiscoverShapes()
        {
            var shapes = Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>().Where(s => s.ElementName().ToUpper().StartsWith("DATA:"));
            Shapes.AddRange(shapes.Select(s => ShapeElement.Create(s.ElementName().Substring(5).Trim(), s, this)).Where(s => s != null));
            var graphicFrames = Slide.Descendants<DocumentFormat.OpenXml.Presentation.GraphicFrame>().Where(s => s.ElementName().ToUpper().StartsWith("DATA:"));
            foreach (var item in graphicFrames)
            {
                string name = item.ElementName().Substring(5).Trim();
                if (item.Graphic.GraphicData.FirstElement<Table>() != null)
                {
                    TableElement table = TableElement.Create(name, item, this);
                    if (table != null)
                    {
                        Shapes.Add(table);
                    }
                }
                else if (item.Graphic.GraphicData.FirstElement<ChartReference>() != null)
                {
                    ChartElement chart = ChartElement.Create(name, item, this);
                    if (chart != null)
                    {
                        Shapes.Add(chart);
                    }
                }
            }
            var paragraphs = Slide.Descendants<Paragraph>();
            foreach (var item in paragraphs)
            {
                if (!item.Ancestors<Table>().Any())
                {
                    if (textTagRegex.Match(item.InnerText).Success)
                    {
                        foreach (Match match in textTagRegex.Matches(item.InnerText))
                        {
                            McKinsey.PowerPointGenerator.Elements.TextElement text = McKinsey.PowerPointGenerator.Elements.TextElement.Create(match.Groups["tag"].Value, item, this);
                            if (text != null)
                            {
                                Shapes.Add(text);
                            }
                        }
                    }
                }
            }
            foreach (ShapeElementBase item in Shapes)
            {
                item.Slide = this;
            }
        }

        public virtual void DiscoverCommands()
        {
            foreach (ShapeElementBase shape in Shapes)
            {
                shape.DiscoverCommands();
            }
        }

        public virtual void FindDataElements(IList<DataElement> data)
        {
            foreach (ShapeElementBase shape in Shapes)
            {
                shape.FindShapeData(data);
            }
        }
    }
}
