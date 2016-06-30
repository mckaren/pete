using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Newtonsoft.Json;

namespace McKinsey.PowerPointGenerator.Core.Data
{
    [JsonObject(MemberSerialization = MemberSerialization.OptIn)]
    [DebuggerDisplay("{Header}, mapped: {MappedHeader}, data: {Data.Count}, hidden: {IsHidden}")]
    public class Row
    {
        [JsonProperty]
        public string Header { get; set; }
        public string MappedHeader { get; set; }
        public List<object> Data { get; private set; }
        public List<object> Legends { get; private set; }
        public DataElement ParentElement { get; set; }
        public bool IsHidden { get; set; }
        public bool IsCore { get; set; }

        public Row()
        {
            Data = new List<object>();
            Legends = new List<object>();
        }

        public string GetHeader()
        {
            if (string.IsNullOrEmpty(MappedHeader))
            {
                return Header;
            }
            return MappedHeader;
        }

        public object this[Index index]
        {
            get
            {
                if (index.Number.HasValue)
                {
                    return this[index.Number.Value];
                }
                else
                {
                    return this[index.Name];
                }
            }
            set
            {
                if (index.Number.HasValue)
                {
                    this[index.Number.Value] = value;
                }
                else
                {
                    this[index.Name] = value;
                }
            }
        }

        public object this[int column]
        {
            get
            {
                try
                {
                    return Data[column];
                }
                catch (ArgumentOutOfRangeException)
                {
                    throw new IndexOutOfRangeException();
                }
            }
            set
            {
                try
                {
                    Data[column] = value;
                }
                catch (ArgumentOutOfRangeException)
                {
                    throw new IndexOutOfRangeException();
                }
            }
        }

        public object this[string columnHeader]
        {
            get
            {
                if (!ParentElement.HasColumnHeaders)
                {
                    throw new ArgumentException("DataElement " + ParentElement.Name + " doesn't have column headers");
                }
                var column = ParentElement.Columns.FirstOrDefault(c => !string.IsNullOrEmpty(c.Header) && c.Header.Equals(columnHeader, StringComparison.OrdinalIgnoreCase));
                if (column != null)
                {
                    int columnIndex = ParentElement.Columns.IndexOf(column);
                    return Data[columnIndex];
                }
                else
                {
                    throw new IndexOutOfRangeException();
                }
            }
            set
            {
                if (!ParentElement.HasColumnHeaders)
                {
                    throw new ArgumentException("DataElement " + ParentElement.Name + " doesn't have column headers");
                }
                var column = ParentElement.Columns.FirstOrDefault(c => !string.IsNullOrEmpty(c.Header) && c.Header.Equals(columnHeader, StringComparison.OrdinalIgnoreCase));
                if (column != null)
                {
                    int columnIndex = ParentElement.Columns.IndexOf(column);
                    Data[columnIndex] = value;
                }
                else
                {
                    throw new IndexOutOfRangeException();
                }

            }
        }

        public Row Clone(bool deepClone = true)
        {
            Row clone = new Row() { Header = Header, ParentElement = ParentElement, MappedHeader = MappedHeader, IsHidden = IsHidden, IsCore = IsCore };
            if (deepClone)
            {
                clone.Data.AddRange(Data);
                clone.Legends.AddRange(Legends);
            }
            return clone;
        }
    }
}
