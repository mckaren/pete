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
using A = DocumentFormat.OpenXml.Drawing;
using McKinsey.PowerPointGenerator.Commands;

namespace McKinsey.PowerPointGenerator.Processing
{
    public class ShapeElementProcessor : IShapeElementProcessor
    {
        private ShapeElement element;

        public void Process(ShapeElementBase shape)
        {
            element = shape as ShapeElement;
            if (element.Data == null)
            {
                return;
            }
            element.Data = element.Data.GetFragmentByIndexes(element.RowIndexes, element.ColumnIndexes);
            element.ProcessCommands(element.Data);

            var visibleCmds = element.CommandsOf<VisibleCommand>();
            bool isVisible = visibleCmds.Count == 0;
            foreach (var vis in visibleCmds)
            {
                isVisible |= vis.IsVisible;
            }

            if (isVisible)
            {
                if (!element.IsContentProtected)
                {
                    var avaiableColumns = element.ColumnIndexes.Where(i => !i.IsHidden && element.Data.HasColumn(i));
                    var avaiableRows = element.RowIndexes.Where(i => !i.IsHidden && element.Data.HasRow(i));
                    Index firstColumn = avaiableColumns.Count() > 0 ? avaiableColumns.First() : new Index(0);
                    Index firstRow = avaiableRows.Count() > 0 ? avaiableRows.First() : new Index(0);
                    object data = element.Data.Data(firstRow, firstColumn);
                    if (data != null)
                    {
                        element.Element.Replace(data.ToString());
                    }
                }

                ShapeElement legend = element.Data.Rows[0].Legends[0] as ShapeElement;
                if (legend != null)
                {
                    A.SolidFill fill = legend.Element.GetFill();
                    A.Outline outline = legend.Element.GetOutline();
                    element.Element.ShapeProperties.ReplaceChild<A.SolidFill>(fill.CloneNode(true), element.Element.ShapeProperties.FirstElement<A.SolidFill>());
                    element.Element.ShapeProperties.ReplaceChild<A.Outline>(outline.CloneNode(true), element.Element.ShapeProperties.FirstElement<A.Outline>());
                }
            }
            else
            {
                element.Element.Remove();
            }
        }
    }
}
