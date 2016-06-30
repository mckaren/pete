using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Elements;
using McKinsey.PowerPointGenerator.Core.Data;

namespace McKinsey.PowerPointGenerator.Commands
{
    /// <summary>
    /// PAGE{rows_per_page, columns_per_page}
    /// Can be used to page charts or tables. The data will be cut into pieces and the whole chart will be repeated with changing data.
    /// Examples:
    /// PAGE{10}
    /// PAGE{10,5}
    /// </summary>
    public class PageCommand : Command
    {
        public static readonly string Name = "PAGE";
        public int RowsPerPage { get; set; }
        public int ColumnsPerPage { get; set; }

        public override void ParseArguments()
        {
            RowsPerPage = Int32.MaxValue;
            ColumnsPerPage = Int32.MaxValue;
            if (!string.IsNullOrEmpty(ArgumentsString))
            {
                string[] parts = ArgumentsString.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length > 0)
                {
                    int rows = 0;
                    if (int.TryParse(parts[0], out rows))
                    {
                        RowsPerPage = rows;
                    }
                }
                if (parts.Length == 2)
                {
                    int cols = 0;
                    if (int.TryParse(parts[1], out cols))
                    {
                        ColumnsPerPage = cols;
                    }
                }
            }
        }

        public override void ApplyToData(DataElement data)
        {
        }
    }
}
