using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Extensions;
using A = DocumentFormat.OpenXml.Drawing;


namespace McKinsey.PowerPointGenerator.Processing
{
    public class TableElementProcessor : IShapeElementProcessor
    {
        private TableElement element;

        public void Process(ShapeElementBase shape)
        {
            element = shape as TableElement;
            if (element.Data == null)
            {
                return;
            }
            element.Data = element.Data.GetFragmentByIndexes(element.RowIndexes, element.ColumnIndexes);
            element.ProcessCommands(element.Data);

            A.GraphicData graphicData = element.TableFrame.Graphic.GraphicData;
            Table table = graphicData.FirstElement<Table>();
            if (table == null)
            {
                return;
            }

            int rowsInTable = table.Elements<TableRow>().Count();
            int columnsInTable = table.Elements<TableRow>().First().Elements<TableCell>().Count();
            if (element.UseColumnHeaders)
            {
                rowsInTable--;
            }
            if (element.UseRowHeaders)
            {
                columnsInTable--;
            }
            if (element.IsFixed)
            {
                element.Data.TrimOrExpand(columnsInTable, rowsInTable, false);
            }
            else
            {
                PrepareTable(table, element.UseRowHeaders, element.UseColumnHeaders);
            }

            int startFromRow = 0;
            if (element.UseColumnHeaders)
            {
                TableRow tableRow = table.Elements<TableRow>().First();
                List<object> rowData = element.Data.Columns.Select(c => (object)c.GetHeader()).ToList();
                FillRowWithHeaders(tableRow);
            }

            for (int rowIndex = 0; rowIndex < element.Data.Rows.Count; rowIndex++)
            {
                int tableRowIndex = element.UseColumnHeaders ? rowIndex + 1 : rowIndex;
                TableRow tableRow = table.Elements<TableRow>().ElementAt(tableRowIndex);
                Row row = element.Data.Rows[rowIndex];

                FillRowWithData(tableRow, row);
                SetCellStyleFromLegend(tableRow, row);
            }
        }

        private void FillRowWithHeaders(TableRow tableRow)
        {
            int tableColumnIndex = 0;
            if (element.UseRowHeaders)
            {
                TableCell tableCell = tableRow.Elements<TableCell>().ElementAt(tableColumnIndex);
                tableCell.ReplaceTextInCellTextBody("");
                tableColumnIndex++;
            }
            foreach (Column column in element.Data.Columns)
            {
                TableCell tableCell = tableRow.Elements<TableCell>().ElementAt(tableColumnIndex);
                tableCell.ReplaceTextInCellTextBody(column.GetHeader());
                tableColumnIndex++;
            }
        }

        private void FillRowWithData(TableRow tableRow, Row row)
        {
            int tableColumnIndex = 0;
            if (element.UseRowHeaders)
            {
                TableCell tableCell = tableRow.Elements<TableCell>().ElementAt(tableColumnIndex);
                tableCell.ReplaceTextInCellTextBody(row.GetHeader());
                tableColumnIndex++;
            }
            for (int dataColumnIndex = 0; dataColumnIndex < row.Data.Count; dataColumnIndex++)
            {
                TableCell tableCell = tableRow.Elements<TableCell>().ElementAt(tableColumnIndex);
                object value = row.Data[dataColumnIndex];
                tableCell.ReplaceTextInCellTextBody(value == null ? "" : value.ToString());
                tableColumnIndex++;
            }
        }

        private void SetCellStyleFromLegend(TableRow tableRow, Row row)
        {
            int tableColumnIndex = 0;
            if (element.UseRowHeaders)
            {
                tableColumnIndex++;
            }

            if (row.Legends.All(l => l == null))
            {
                return;
            }
            for (int dataColumnIndex = 0; dataColumnIndex < row.Data.Count; dataColumnIndex++)
            {
                if (row.Legends[dataColumnIndex] != null)
                {
                    TableCell tableCell = tableRow.Elements<TableCell>().ElementAt(tableColumnIndex);
                    ShapeElement legend = row.Legends[dataColumnIndex] as ShapeElement;
                    if (legend != null)
                    {
                        A.SolidFill fill = legend.Element.GetFill();
                        A.Outline outline = legend.Element.GetOutline();
                        A.SolidFill outlineSolidFill = outline.FirstElement<A.SolidFill>();
                        if (tableCell.TableCellProperties.Elements<A.SolidFill>().Count() == 0)
                        {
                            tableCell.TableCellProperties.AppendChild<A.SolidFill>(fill.CloneNode(true) as A.SolidFill);
                        }
                        else
                        {
                            tableCell.TableCellProperties.ReplaceChild<A.SolidFill>(fill.CloneNode(true), tableCell.TableCellProperties.FirstElement<A.SolidFill>());
                        }
                        if (outlineSolidFill != null)
                        {
                            if (tableCell.TableCellProperties.LeftBorderLineProperties != null)
                            {
                                tableCell.TableCellProperties.LeftBorderLineProperties.RemoveAllChildren<NoFill>();
                                tableCell.TableCellProperties.LeftBorderLineProperties.RemoveAllChildren<A.SolidFill>();
                                tableCell.TableCellProperties.LeftBorderLineProperties.Append(outlineSolidFill.CloneNode(true));
                            }
                            if (tableCell.TableCellProperties.TopBorderLineProperties != null)
                            {
                                tableCell.TableCellProperties.TopBorderLineProperties.RemoveAllChildren<NoFill>();
                                tableCell.TableCellProperties.TopBorderLineProperties.RemoveAllChildren<A.SolidFill>();
                                tableCell.TableCellProperties.TopBorderLineProperties.Append(outlineSolidFill.CloneNode(true));
                            }
                            if (tableCell.TableCellProperties.RightBorderLineProperties != null)
                            {
                                tableCell.TableCellProperties.RightBorderLineProperties.RemoveAllChildren<NoFill>();
                                tableCell.TableCellProperties.RightBorderLineProperties.RemoveAllChildren<A.SolidFill>();
                                tableCell.TableCellProperties.RightBorderLineProperties.Append(outlineSolidFill.CloneNode(true));
                            }
                            if (tableCell.TableCellProperties.BottomBorderLineProperties != null)
                            {
                                tableCell.TableCellProperties.BottomBorderLineProperties.RemoveAllChildren<NoFill>();
                                tableCell.TableCellProperties.BottomBorderLineProperties.RemoveAllChildren<A.SolidFill>();
                                tableCell.TableCellProperties.BottomBorderLineProperties.Append(outlineSolidFill.CloneNode(true));
                            }
                        }
                        
                        TextCharacterPropertiesType prop = legend.Element.GetRunProperties();
                        foreach (Paragraph paragraph in tableCell.TextBody.Elements<Paragraph>())
                        {
                            foreach (Run run in paragraph.Elements<Run>())
                            {
                                run.RunProperties.RemoveAllChildren();
                                foreach (var item in prop.ChildElements)
                                {
                                    run.RunProperties.AppendChild(item.CloneNode(true));
                                }
                            }
                            EndParagraphRunProperties endParagraphProperties = paragraph.FirstElement<EndParagraphRunProperties>();
                            if (endParagraphProperties != null)
                            {
                                endParagraphProperties.RemoveAllChildren();
                                foreach (var item in prop.ChildElements)
                                {
                                    endParagraphProperties.AppendChild(item.CloneNode(true));
                                }
                            }
                        }
                    }
                }
                tableColumnIndex++;
            }
        }

        private void PrepareTable(Table table, bool useRowHeaders, bool useColumnHeaders)
        {
            int tableRowsCount = table.Elements<TableRow>().Count();
            int dataRowsCount = useColumnHeaders ? element.Data.Rows.Count + 1 : element.Data.Rows.Count;
            if (tableRowsCount < dataRowsCount)
            {
                for (int i = tableRowsCount; i < dataRowsCount; i++)
                {
                    TableRow lastRowClone = table.Elements<TableRow>().Last().CloneNode(true) as TableRow;
                    table.Append(lastRowClone);
                }
            }
            else
            {
                for (int i = dataRowsCount; i < tableRowsCount; i++)
                {
                    table.RemoveChild<TableRow>(table.Elements<TableRow>().Last());
                }
            }

            int tableColumnsCount = table.Elements<TableRow>().First().Elements<TableCell>().Count();
            int dataColumnsCount = useRowHeaders ? element.Data.Columns.Count + 1 : element.Data.Columns.Count;

            var tableGrid = table.FirstElement<TableGrid>();
            long totalWidthBefore = tableGrid.Elements<GridColumn>().Sum(c => c.Width);
            if (tableColumnsCount < dataColumnsCount)
            {
                foreach (TableRow tableRow in table.Elements<TableRow>())
                {
                    for (int i = tableColumnsCount; i < dataColumnsCount; i++)
                    {
                        TableCell lastCellClone = tableRow.Elements<TableCell>().Last().CloneNode(true) as TableCell;
                        tableRow.Append(lastCellClone);
                    }

                }
                if (tableGrid != null)
                {
                    for (int i = tableColumnsCount; i < dataColumnsCount; i++)
                    {
                        GridColumn lastColumnClone = tableGrid.Elements<GridColumn>().Last().CloneNode(true) as GridColumn;
                        tableGrid.Append(lastColumnClone);
                    }
                }
            }
            else
            {
                foreach (TableRow tableRow in table.Elements<TableRow>())
                {
                    for (int i = dataColumnsCount; i < tableColumnsCount; i++)
                    {
                        tableRow.RemoveChild<TableCell>(tableRow.Elements<TableCell>().Last());
                    }
                }
                if (tableGrid != null)
                {
                    for (int i = dataColumnsCount; i < tableColumnsCount; i++)
                    {
                        tableGrid.RemoveChild<GridColumn>(tableGrid.Elements<GridColumn>().Last());
                    }
                }
            }
            //if (!element.IsDynamic)
            //{
                long totalWidthAfter = tableGrid.Elements<GridColumn>().Sum(c => c.Width);
                if (totalWidthBefore != totalWidthAfter)
                {
                    var scale = (decimal)totalWidthBefore / (decimal)totalWidthAfter;
                    foreach (GridColumn gridColumn in tableGrid.Elements<GridColumn>())
                    {
                        gridColumn.Width = (long)(gridColumn.Width * scale);
                    }
                }
            //}
        }
    }
}
