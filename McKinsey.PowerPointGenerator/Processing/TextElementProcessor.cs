using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Drawing;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.Practices.Unity;
using McKinsey.PowerPointGenerator.Extensions;

namespace McKinsey.PowerPointGenerator.Processing
{
    public class TextElementProcessor : IShapeElementProcessor
    {
        private TextElement element;

        public void Process(ShapeElementBase shape)
        {
            element = shape as TextElement;
            if (element.Data == null)
            {
                return;
            }
            element.Data = element.Data.GetFragmentByIndexes(element.RowIndexes, element.ColumnIndexes);
            element.ProcessCommands(element.Data);

            if (element != null)
            {
                object data = element.Data.Data(0, 0);
                if (data != null)
                {
                    element.Paragraph.Replace("#" + element.FullName + "#", data.ToString());
                }
            }
        }
    }
}
