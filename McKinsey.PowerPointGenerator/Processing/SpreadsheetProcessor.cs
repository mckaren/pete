using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using SpreadsheetGear;

namespace McKinsey.PowerPointGenerator.Processing
{
    public static class SpreadsheetProcessor
    {
        public static List<ChartSeriesElement> InsertData(DataElement data, Stream inputStream, Stream outputStream)
        {
            List<ChartSeriesElement> series = new List<ChartSeriesElement>();
            IWorkbookSet workbookSet = Factory.GetWorkbookSet();
            IWorkbook workbook = workbookSet.Workbooks.OpenFromStream(inputStream);
            IWorksheet worksheet = workbook.Worksheets[0];
            worksheet.UsedRange.Clear();
            int rowNo = 0, columnNo = 0;
            string axisDataFormula = null;
            rowNo++;
            axisDataFormula = worksheet.Cells[1, 0, data.Rows.Count, 0].Address;
            foreach (Row row in data.Rows)
            {
                worksheet.Cells[rowNo, columnNo].Value = row.Header;
                rowNo++;
            }
            columnNo++;
            rowNo = 0;
            //}
            var columnsToInclude = data.Columns.ToList();
            for (int colInd = 0; colInd < columnsToInclude.Count; colInd++)
            {
                string seriesTextFormula = null;
                string valuesFormula = null;
                Index columnIndex = null;
                worksheet.Cells[rowNo, columnNo].Value = columnsToInclude[colInd].Header;
                rowNo++;
                seriesTextFormula = worksheet.Cells[0, columnNo].Address;
                valuesFormula = worksheet.Cells[1, columnNo, data.Rows.Count, columnNo].Address;
                if (string.IsNullOrEmpty((columnsToInclude[colInd].Header)))
                {
                    columnIndex = new Index(colInd) { IsCore = columnsToInclude[colInd].IsCore, IsHidden = columnsToInclude[colInd].IsHidden };
                }
                else
                {
                    columnIndex = new Index(columnsToInclude[colInd].Header) { IsCore = columnsToInclude[colInd].IsCore, IsHidden = columnsToInclude[colInd].IsHidden };
                }
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    worksheet.Cells[rowNo, columnNo].Value = columnsToInclude[colInd].Data[i];
                    rowNo++;
                }
                series.Add(new ChartSeriesElement { CategoryAxisDataAddress = worksheet.Name + "!" + axisDataFormula, SeriesTextAddress = worksheet.Name + "!" + seriesTextFormula, ValuesAddress = worksheet.Name + "!" + valuesFormula, ColumnIndex = columnIndex });
                columnNo++;
                rowNo = 0;
            }
            workbook.SaveToStream(outputStream, FileFormat.OpenXMLWorkbook);
            return series;
        }
    }
}
