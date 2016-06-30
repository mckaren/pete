using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Core.Data;

namespace McKinsey.PowerPointGenerator.Elements
{
    [DebuggerDisplay("Column: {ColumnIndex}, text: {SeriesTextAddress}, values: {ValuesAddress}, categories: {CategoryAxisDataAddress}")]
    public class ChartSeriesElement
    {
        public string SeriesTextAddress { get; set; }
        public string ValuesAddress { get; set; }
        public string CategoryAxisDataAddress { get; set; }
        public Index ColumnIndex { get; set; }
        public ChartSeriesElement YValues { get; set; }
        public ChartSeriesElement MinusErrorBar { get; set; }
        public ChartSeriesElement PlusErrorBar { get; set; }
    }
}
