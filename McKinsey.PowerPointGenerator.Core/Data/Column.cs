using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Newtonsoft.Json;

namespace McKinsey.PowerPointGenerator.Core.Data
{
    [JsonObject(MemberSerialization = MemberSerialization.OptIn)]
    [DebuggerDisplay("{Header}, mapped: {MappedHeader}, data: {Data.Count}, hidden: {IsHidden}, calculated: {IsCalculated}")]
    public class Column
    {
        [JsonProperty]
        public string Header { get; set; }
        public string MappedHeader { get; set; }
        [JsonProperty]
        public List<object> Data { get; private set; }
        public List<object> Legends { get; private set; }
        public DataElement ParentElement { get; set; }
        public bool IsCalculated { get; set; }
        public bool IsHidden { get; set; }
        public bool IsCore { get; set; }

        public Column()
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

        public object this[int row]
        {
            get
            {
                try
                {
                    return Data[row];
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
                    Data[row] = value;
                }
                catch (ArgumentOutOfRangeException)
                {
                    throw new IndexOutOfRangeException();
                }
            }
        }

        public object this[string rowHeader]
        {
            get
            {
                if (!ParentElement.HasRowHeaders)
                {
                    throw new ArgumentException("DataElement " + ParentElement.Name + " doesn't have row headers");
                }
                var row = ParentElement.Rows.FirstOrDefault(c => !string.IsNullOrEmpty(c.Header) && c.Header.Equals(rowHeader, StringComparison.OrdinalIgnoreCase));
                if (row != null)
                {
                    int columnIndex = ParentElement.Rows.IndexOf(row);
                    return Data[columnIndex];
                }
                else
                {
                    throw new IndexOutOfRangeException();
                }
            }
            set
            {
                if (!ParentElement.HasRowHeaders)
                {
                    throw new ArgumentException("DataElement " + ParentElement.Name + " doesn't have row headers");
                }
                var row = ParentElement.Rows.FirstOrDefault(c => !string.IsNullOrEmpty(c.Header) && c.Header.Equals(rowHeader, StringComparison.OrdinalIgnoreCase));
                if (row != null)
                {
                    int columnIndex = ParentElement.Rows.IndexOf(row);
                    Data[columnIndex] = value;
                }
                else
                {
                    throw new IndexOutOfRangeException();
                }

            }
        }

        public Column Clone(bool deepClone = true)
        {
            Column clone = new Column() { Header = Header, ParentElement = ParentElement, MappedHeader = MappedHeader, IsHidden = IsHidden, IsCore = IsCore };
            if (deepClone)
            {
                clone.Data.AddRange(Data);
                clone.Legends.AddRange(Legends);
            }
            return clone;
        }
    }
}
