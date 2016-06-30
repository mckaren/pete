using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace McKinsey.PowerPointGenerator.Extensions
{
    public static class TableCellExtensions
    {
        /// <summary>
        /// Searching for specified template in the table cell and replacing it
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="cell">The table cell.</param>
        /// <returns>true if replacement found</returns>
        /// <param name="templateString"></param>
        public static void ReplaceTextInCellTextBody(this TableCell cell, string value, SolidFill fill = null)
        {
            Paragraph paragraph = cell.TextBody.Elements<Paragraph>().First();
            while (cell.TextBody.Elements<Paragraph>().Count() > 1)
            {
                cell.TextBody.RemoveChild<Paragraph>(cell.TextBody.Elements<Paragraph>().Last());
            }
            paragraph.Replace(value);
        }
    }
}
