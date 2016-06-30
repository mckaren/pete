using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using A = DocumentFormat.OpenXml.Drawing;

namespace McKinsey.PowerPointGenerator.Extensions
{
    public static class ShapeExtensions
    {
        /// <summary>
        /// Gets the element name
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns>The name</returns>
        public static string ElementName(this DocumentFormat.OpenXml.Presentation.Shape shape)
        {
            return shape.NonVisualShapeProperties.NonVisualDrawingProperties.Name.Value;
        }

        /// <summary>
        /// Gets the element name
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <returns>The name</returns>
        public static string ElementName(this DocumentFormat.OpenXml.Presentation.GraphicFrame shape)
        {
            return shape.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name.Value;
        }

        /// <summary>
        /// Replaces the content of the shape.
        /// </summary>
        /// <param name="shape">The shape.</param>
        /// <param name="newValue">The new value.</param>
        /// <param name="fill">The fill.</param>
        public static void Replace(this DocumentFormat.OpenXml.Presentation.Shape shape, string newValue, SolidFill fill = null)
        {
            if (shape.InnerText.Contains("##"))
            {
                foreach (Paragraph paragraph in shape.TextBody.Elements<Paragraph>())
                {
                    paragraph.Replace("##", newValue, fill);
                }
            }
            else
            {
                if (shape.TextBody.Elements<Paragraph>().Any())
                {
                    shape.TextBody.Elements<Paragraph>().First().Replace(newValue, fill);
                }
            }
        }

        public static A.SolidFill GetFill(this DocumentFormat.OpenXml.Presentation.Shape shape)
        {
            A.SolidFill fill = shape.ShapeProperties.FirstElement<A.SolidFill>();
            if (fill != null)
            {
                return fill.CloneNode(true) as A.SolidFill;
            }
            return null;
        }

        public static A.Outline GetOutline(this DocumentFormat.OpenXml.Presentation.Shape shape)
        {
            A.Outline outline = shape.ShapeProperties.FirstElement<A.Outline>();
            if (outline != null)
            {
                return outline.CloneNode(true) as A.Outline;
            }
            return null;
        }

        public static TextCharacterPropertiesType GetRunProperties(this DocumentFormat.OpenXml.Presentation.Shape shape)
        {
            Paragraph paragraph = shape.TextBody.FirstElement<Paragraph>();
            if (paragraph != null)
            {
                Run run = paragraph.FirstElement<Run>();
                if (run != null)
                {
                    return run.RunProperties.CloneNode(true) as RunProperties;
                }
                EndParagraphRunProperties endProp = paragraph.FirstElement<EndParagraphRunProperties>();
                return endProp.CloneNode(true) as EndParagraphRunProperties;
            }
            return null;
        }
    }
}