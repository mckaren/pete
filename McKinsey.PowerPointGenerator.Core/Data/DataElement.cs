using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Serialization;
using Newtonsoft.Json;

namespace McKinsey.PowerPointGenerator.Core.Data
{
    [JsonObject(MemberSerialization = MemberSerialization.OptIn)]
    [DebuggerDisplay("{Name}")]
    public class DataElement
    {
        public bool HasColumnHeaders { get; set; }
        public bool HasRowHeaders { get; set; }
        [JsonProperty]
        public List<Column> Columns { get; private set; }
        [JsonProperty]
        public List<Row> Rows { get; private set; }
        [JsonProperty]
        public virtual string Name { get; set; }
        public bool IsTransposed { get; set; }

        public DataElement()
        {
            Columns = new List<Column>();
            Rows = new List<Row>();
        }

        public DataElement GetFragmentByIndexes(List<Index> rowIndexes, List<Index> columnIndexes)
        {
            if (rowIndexes.Count == 0 && columnIndexes.Count == 0 || (rowIndexes.Count == 1 && rowIndexes[0].IsAll && columnIndexes.Count == 1 && columnIndexes[0].IsAll))
            {
                return this;
            }

            DataElement fragment = new DataElement { Name = Name, HasColumnHeaders = HasColumnHeaders, HasRowHeaders = HasRowHeaders };

            List<Row> rowsToTake;
            bool takeAllRows = rowIndexes.Count == 1 && rowIndexes[0].IsAll; // rowIndexes.Any(i => i.IsAll) || rowIndexes.Where(i => HasRow(i)).Count() == 0;
            bool takeAllColumns = columnIndexes.Count == 1 && columnIndexes[0].IsAll;  //columnIndexes.Any(i => i.IsAll) || columnIndexes.Where(i => HasColumn(i)).Count() == 0;
            if (rowIndexes.Count == 0 || takeAllRows)
            {
                rowsToTake = Rows;
            }
            else
            {
                rowsToTake = new List<Row>();
                foreach (Index rowIndex in rowIndexes.Where(r => HasRow(r)))
                {
                    Row row = Row(rowIndex);
                    row.IsCore = rowIndex.IsCore;
                    row.IsHidden = rowIndex.IsHidden;
                    rowsToTake.Add(row);
                }
            }

            foreach (Row row in rowsToTake)
            {
                Row newRow;
                if (columnIndexes.Count == 0 || takeAllColumns)
                {
                    newRow = row.Clone();
                }
                else
                {
                    newRow = new Row() { Header = row.Header, MappedHeader = row.MappedHeader, IsHidden = row.IsHidden };
                    foreach (Index columnIndex in columnIndexes)
                    {
                        if (HasColumn(columnIndex))
                        {
                            newRow.Data.Add(row[columnIndex]);
                        }
                    }
                }
                fragment.Rows.Add(newRow);
            }
            if (columnIndexes.Count == 0 || takeAllColumns)
            {
                foreach (Column column in Columns)
                {
                    fragment.Columns.Add(new Column { Header = column.Header, MappedHeader = column.MappedHeader, IsHidden = column.IsHidden });
                }
            }
            else
            {
                foreach (Index columnIndex in columnIndexes.Where(c => HasColumn(c)))
                {
                    Column column = Column(columnIndex);
                    fragment.Columns.Add(new Column { Header = column.Header, MappedHeader = column.MappedHeader, IsHidden = columnIndex.IsHidden, IsCore = columnIndex.IsCore });
                }
            }
            fragment.Normalize();
            return fragment;
        }

        public void TrimHiddenRowsAndColumns()
        {
            for (int i = Rows.Count - 1; i >= 0; i--)
            {
                if (Rows[i].IsHidden)
                {
                    Rows.RemoveAt(i);
                    foreach (Column column in Columns)
                    {
                        column.Data.RemoveAt(i);
                    }
                }
            }
            for (int i = Columns.Count - 1; i >= 0; i--)
            {
                if (Columns[i].IsHidden)
                {
                    Columns.RemoveAt(i);
                    foreach (Row row in Rows)
                    {
                        row.Data.RemoveAt(i);
                    }
                }
            }
        }


        public object Data(Index rowIndex, Index columIndex)
        {
            return Row(rowIndex)[columIndex];
        }

        public object Data(int rowIndex, int columIndex)
        {
            return Rows[rowIndex].Data[columIndex];
        }

        public object Data(Index rowIndex, Index columIndex, object newValue)
        {
            Row(rowIndex)[columIndex] = newValue;
            Column(columIndex)[rowIndex] = newValue;
            return newValue;
        }

        public bool HasRow(Index index)
        {
            if (Rows.Any(r => !string.IsNullOrEmpty(r.Header) && r.Header.Equals(index.Name, StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }
            if (index.Number.HasValue)
            {
                if (index.Number.Value < Rows.Count)
                {
                    return true;
                }
            }
            return false;
        }

        public bool HasColumn(Index index)
        {
            if (Columns.Any(r => !string.IsNullOrEmpty(r.Header) && r.Header.Equals(index.Name, StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }
            if (index.Number.HasValue)
            {
                if (index.Number.Value < Columns.Count)
                {
                    return true;
                }
            }
            return false;
        }

        public Row Row(Index index)
        {
            if (index.Number.HasValue)
            {
                return Row(index.Number.Value);
            }
            else
            {
                return Row(index.Name);
            }
        }

        public Row Row(int row)
        {
            try
            {
                return Rows[row];
            }
            catch (ArgumentOutOfRangeException)
            {
                throw new IndexOutOfRangeException();
            }
        }

        public Row Row(string rowHeader)
        {
            if (!HasRowHeaders)
            {
                throw new ArgumentException("DataElement " + Name + " doesn't have row headers");
            }

            var row = Rows.FirstOrDefault(r => r.Header.Equals(rowHeader, StringComparison.OrdinalIgnoreCase));
            if (row == null)
            {
                throw new IndexOutOfRangeException();
            }
            return row;
        }

        public Column Column(Index index)
        {
            if (index.Number.HasValue)
            {
                return Column(index.Number.Value);
            }
            else
            {
                return Column(index.Name);
            }
        }

        public Column Column(int column)
        {
            try
            {
                return Columns[column];
            }
            catch (ArgumentOutOfRangeException)
            {
                throw new IndexOutOfRangeException();
            }
        }

        public Column Column(string columnHeader)
        {
            if (!HasColumnHeaders)
            {
                throw new ArgumentException("DataElement " + Name + " doesn't have column headers");
            }
            var column = Columns.FirstOrDefault(r => !string.IsNullOrEmpty(r.Header) && r.Header.Equals(columnHeader, StringComparison.OrdinalIgnoreCase));
            if (column == null)
            {
                throw new IndexOutOfRangeException();
            }
            return column;
        }

        public virtual DataElement Clone()
        {
            DataElement clone = new DataElement() { Name = Name, HasColumnHeaders = HasColumnHeaders, HasRowHeaders = HasRowHeaders, IsTransposed = IsTransposed };
            clone.Rows.AddRange(Rows.Select(r => r.Clone()));
            clone.Columns.AddRange(Columns.Select(c => c.Clone()));
            return clone;
        }

        public void Normalize()
        {
            bool copyFromRowsToColumns = Columns.Count == 0 || Columns[0].Data.Count == 0;
            bool copyFromColumnsToRows = Rows.Count == 0 || Rows[0].Data.Count == 0;

            if (copyFromRowsToColumns)
            {
                foreach (Row row in Rows)
                {
                    if (row.Data.Count > row.Legends.Count)
                    {
                        row.Legends.AddRange(Enumerable.Repeat<object>(null, row.Data.Count - row.Legends.Count).ToList());
                    }
                    if (row.Data.Count < row.Legends.Count)
                    {
                        row.Legends.RemoveRange(row.Data.Count, row.Legends.Count - row.Data.Count);
                    }
                }
            }
            if (copyFromColumnsToRows)
            {
                foreach (Column column in Columns)
                {
                    if (column.Data.Count > column.Legends.Count)
                    {
                        column.Legends.AddRange(Enumerable.Repeat<object>(null, column.Data.Count - column.Legends.Count).ToList());
                    }
                    if (column.Data.Count < column.Legends.Count)
                    {
                        column.Legends.RemoveRange(column.Data.Count, column.Legends.Count - column.Data.Count);
                    }
                }
            }

            foreach (Row row in Rows)
            {
                row.ParentElement = this;
                if (copyFromRowsToColumns)
                {
                    bool addColumn = false;
                    if (Columns.Count < row.Data.Count)
                    {
                        addColumn = true;
                    }
                    for (int i = 0; i < row.Data.Count; i++)
                    {
                        if (addColumn)
                        {
                            Columns.Add(new Column());
                        }
                        Columns[i].Data.Add(row.Data[i]);
                        Columns[i].Legends.Add(row.Legends[i]);
                    }
                }
            }
            foreach (Column column in Columns)
            {
                column.ParentElement = this;
                if (copyFromColumnsToRows)
                {
                    bool addRow = false;
                    if (Rows.Count < column.Data.Count)
                    {
                        addRow = true;
                    }
                    for (int i = 0; i < column.Data.Count; i++)
                    {
                        if (addRow)
                        {
                            Rows.Add(new Row());
                        }
                        Rows[i].Data.Add(column.Data[i]);
                        Rows[i].Legends.Add(column.Legends[i]);
                    }
                }
            }
        }

        public void TrimOrExpand(int columnsLimit, int rowsLimit, bool fillWhenExpand = true)
        {
            if (!fillWhenExpand)
            {
                if (columnsLimit > Columns.Count)
                {
                    columnsLimit = Columns.Count;
                }
                if (rowsLimit > Rows.Count)
                {
                    rowsLimit = Rows.Count;
                }
            }
            if (columnsLimit == Columns.Count && rowsLimit == Rows.Count)
            {
                return;
            }

            var temp = this.Clone();
            Columns.Clear();
            Rows.Clear();
            for (int i = 0; i < rowsLimit; i++)
            {
                Row newRow = new Row();
                for (int k = 0; k < columnsLimit; k++)
                {
                    newRow.Data.Add(null);
                    newRow.Legends.Add(null);
                }
                Rows.Add(newRow);
            }
            for (int k = 0; k < columnsLimit; k++)
            {
                Columns.Add(new Column());
            }

            int columnRange = Math.Min(columnsLimit, temp.Columns.Count);
            int rowRange = Math.Min(rowsLimit, temp.Rows.Count);

            for (int i = 0; i < rowRange; i++)
            {
                Rows[i].Header = temp.Rows[i].Header;
                Rows[i].IsHidden = temp.Rows[i].IsHidden;
                Rows[i].MappedHeader = temp.Rows[i].MappedHeader;
                for (int k = 0; k < columnRange; k++)
                {
                    Rows[i].Data[k] = temp.Rows[i].Data[k];
                    Rows[i].Legends[k] = temp.Rows[i].Legends[k];
                }
            }
            for (int k = 0; k < columnRange; k++)
            {
                Columns[k].Header = temp.Columns[k].Header;
                Columns[k].IsHidden = temp.Columns[k].IsHidden;
                Columns[k].MappedHeader = temp.Columns[k].MappedHeader;
                Columns[k].IsCalculated = temp.Columns[k].IsCalculated;
            }
            Normalize();
        }

        public virtual void MergeWith(DataElement mergedElement)
        {
            int initialColumnsCount = Columns.Count;
            int initialRowsCount = Rows.Count;
            for (int rowInd = 0; rowInd < mergedElement.Rows.Count; rowInd++)
            {
                if (string.IsNullOrEmpty(mergedElement.Rows[rowInd].Header) || !Rows.Any(r => r.Header.Equals(mergedElement.Rows[rowInd].Header, StringComparison.OrdinalIgnoreCase)))
                {
                    Row newRow = mergedElement.Rows[rowInd].Clone();
                    newRow.Data.InsertRange(0, Enumerable.Repeat<object>(null, initialColumnsCount));
                    Rows.Add(newRow);
                }
            }
            for (int colInd = 0; colInd < mergedElement.Columns.Count; colInd++)
            {
                if (string.IsNullOrEmpty(mergedElement.Columns[colInd].Header) || !Columns.Any(r => r.Header.Equals(mergedElement.Columns[colInd].Header, StringComparison.OrdinalIgnoreCase)))
                {
                    Column newColumn = mergedElement.Columns[colInd].Clone(false);
                    Columns.Add(newColumn);
                    foreach (Row row in Rows)
                    {
                        row.Data.Add(null);
                    }
                }
            }

            for (int rowInd = 0; rowInd < mergedElement.Rows.Count; rowInd++)
            {
                Row row = Rows.FirstOrDefault(r => !string.IsNullOrEmpty(r.Header) && r.Header.Equals(mergedElement.Rows[rowInd].Header, StringComparison.OrdinalIgnoreCase));
                if (row == null)
                {
                    row = Rows[initialRowsCount + rowInd];
                }
                for (int colInd = 0; colInd < mergedElement.Columns.Count; colInd++)
                {
                    if (mergedElement.Rows[rowInd].Data[colInd] != null)
                    {
                        int dataIndex = initialColumnsCount + colInd;
                        Column column = Columns.FirstOrDefault(r => !string.IsNullOrEmpty(r.Header) && r.Header.Equals(mergedElement.Columns[colInd].Header, StringComparison.OrdinalIgnoreCase));
                        if (column != null)
                        {
                            dataIndex = Columns.IndexOf(column);
                        }
                        if (row.Data[dataIndex] == null)
                        {
                            row.Data[dataIndex] = mergedElement.Rows[rowInd].Data[colInd];
                        }
                    }
                }
            }
            foreach (var item in Columns)
            {
                item.Data.Clear();
            }

            Normalize();
        }

        [OnDeserialized]
        internal void OnDeserializedMethod(StreamingContext context)
        {
            Normalize();
        }
    }
}