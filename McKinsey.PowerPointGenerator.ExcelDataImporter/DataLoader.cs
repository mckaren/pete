using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Core.Data;
using NLog;
using SpreadsheetGear;

namespace McKinsey.PowerPointGenerator.ExcelDataImporter
{
    public class DataLoader
    {
        private IWorkbook workbook;
        private int workbookNamesCount;

        public void Import(Stream stream, Action<int, int> progressCallback = null)
        {
            IWorkbookSet workbookSet = Factory.GetWorkbookSet();
            workbookSet.CalculationOnDemand = true;
            workbook = workbookSet.Workbooks.OpenFromStream(stream);
            workbookSet.CalculateFullRebuild();
            workbookNamesCount = workbook.Names.Count;
            Logger logger = LogManager.GetLogger("Generator");
            logger.Debug("Found {0} data elements", workbookNamesCount);
            if (progressCallback != null)
            {
                progressCallback(0, workbookNamesCount);
            }
        }

        public IList<DataElement> LoadData(Action<int, int> progressCallback = null)
        {
            workbook.WorkbookSet.GetLock();
            IList<DataElement> result = new List<DataElement>();
            try
            {
                for (int i = 0; i < workbookNamesCount; i++)
                {
                    if (workbook.Names[i].Visible)
                    {
                        IRange range = workbook.Names[i].RefersToRange;
                        if (range != null)
                        {
                            if (progressCallback != null)
                            {
                                progressCallback(i, workbookNamesCount);
                            }
                            range = range.Intersect(range.Worksheet.UsedRange);
                            result.Add(CreateDataElementFromRange(workbook.Names[i].Name, range));
                        }
                    }
                }
            }
            finally
            {
                workbook.WorkbookSet.ReleaseLock();
            }
            return result;
        }

        private DataElement CreateDataElementFromRange(string name, IRange range)
        {
            try
            {
                if (range.RowCount == 1 && range.ColumnCount == 1)
                {
                    return CreateAtomicDataElement(name, GetValue(range[0, 0]));
                }

                bool hasColumnHeaders = range.RowCount > 1 && range.ColumnCount > 1;
                bool hasRowHeaders = range.RowCount > 1 && range.ColumnCount > 1;
                DataElement element = new DataElement { Name = name, HasColumnHeaders = hasColumnHeaders, HasRowHeaders = hasRowHeaders };
                for (int rowIndex = 0; rowIndex < range.RowCount; rowIndex++)
                {
                    var dataRow = (IRange)range[rowIndex, 0, rowIndex, range.ColumnCount - 1];
                    if (rowIndex == 0 && hasColumnHeaders)
                    {
                        for (int columnIndex = 1; columnIndex < dataRow.ColumnCount; columnIndex++)
                        {
                            element.Columns.Add(new Column { Header = dataRow[0, columnIndex].Text });
                        }
                    }
                    else
                    {
                        int startRow = hasRowHeaders ? 1 : 0;
                        Row row = new Row();
                        if (hasRowHeaders)
                        {
                            row.Header = dataRow[0, 0].Text;
                        }
                        for (int columnIndex = startRow; columnIndex < dataRow.ColumnCount; columnIndex++)
                        {
                            if (!string.IsNullOrEmpty(dataRow[0, columnIndex].Text))
                            {
                                row.Data.Add(GetValue(dataRow[0, columnIndex]));
                            }
                            else
                            {
                                row.Data.Add(null);
                            }
                        }
                        element.Rows.Add(row);
                    }
                }
                element.Normalize();
                return element;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private DataElement CreateAtomicDataElement(string name, object value)
        {
            DataElement element = new DataElement { Name = name, HasColumnHeaders = false, HasRowHeaders = false };
            Row row = new Row { ParentElement = element };
            Column column = new Column { ParentElement = element };
            row.Data.Add(value);
            row.Legends.Add(null);
            column.Data.Add(value);
            column.Legends.Add(null);
            element.Rows.Add(row);
            element.Columns.Add(column);
            return element;
        }

        private object GetValue(IRange range)
        {
            if (string.IsNullOrEmpty(range[0, 0].Text))
            {
                return null;
            }
            try
            {
                switch (range.NumberFormatType)
                {
                    case NumberFormatType.Currency:
                    case NumberFormatType.Number:
                    case NumberFormatType.Percent:
                    case NumberFormatType.Scientific:
                    case NumberFormatType.Fraction:
                        return Convert.ToDouble(range[0, 0].Value);
                        break;
                    case NumberFormatType.Time:
                    case NumberFormatType.Date:
                    case NumberFormatType.DateTime:
                        double value = Convert.ToDouble(range[0, 0].Value);
                        return range.Worksheet.Workbook.NumberToDateTime(value);
                        break;
                    case NumberFormatType.None:
                        return range[0, 0].Value;
                        break;
                    case NumberFormatType.General:
                    case NumberFormatType.Text:
                        return range[0, 0].Text;
                        break;
                }
            }
            catch (Exception ex)
            {
                return range[0, 0].Value;
            }
            return range[0, 0].Value;
        }
    }
}
