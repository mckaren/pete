using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Elements;
using NLog;

namespace McKinsey.PowerPointGenerator.Processing
{
    public static class DataElementProcessor
    {
        private static Regex textTagRegex = new Regex(@"#(?<tag>.*?(?:\:\{.*?\})?)#", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static Regex tagParseRegex = new Regex(@"^(?<name>[\w_]+)(?:\[(?<column>.*?)\])?(?:\[(?<row>.*?)\])?(?:(?:\:+)(?<format>.*))?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static void Process(IList<DataElement> data)
        {
            foreach (DataElement dataElement in data)
            {
                ProcessDataElement(dataElement, data);
            }
        }

        public static bool ProcessDataElement(DataElement dataElement, IList<DataElement> dataSet)
        {
            bool replaced = false;
            bool dataReplaced = false;
            foreach (Row row in dataElement.Rows)
            {
                string rowHeaderReplacement = ParseAndReplace(row.Header, dataSet);
                if (!string.IsNullOrEmpty(rowHeaderReplacement) && !rowHeaderReplacement.Equals(row.Header))
                {
                    row.MappedHeader = rowHeaderReplacement;
                    replaced = true;
                }
                for (int i = 0; i < row.Data.Count; i++)
                {
                    if (row.Data[i] is string)
                    {
                        string dataReplacement = ParseAndReplace((string)row.Data[i], dataSet);
                        if (!string.IsNullOrEmpty(dataReplacement) && !dataReplacement.Equals((string)row.Data[i]))
                        {
                            row.Data[i] = dataReplacement;
                            replaced = true;
                            dataReplaced = true;
                        }
                    }
                }
            }
            foreach (Column column in dataElement.Columns)
            {
                string columnHeaderReplacement = ParseAndReplace(column.Header, dataSet);
                if (!string.IsNullOrEmpty(columnHeaderReplacement) && !columnHeaderReplacement.Equals(column.Header))
                {
                    column.MappedHeader = columnHeaderReplacement;
                    replaced = true;
                }
                if (dataReplaced)
                {
                    column.Data.Clear();
                }
            }
            if (dataReplaced)
            {
                dataElement.Normalize();
            }
            return replaced;
        }

        internal static string ParseAndReplace(string data, IList<DataElement> dataSet, int depth = 0)
        {
            if (depth > 10)
            {
                throw new InvalidOperationException("There are circular references in data tags. The system stopped processing after 10 iterations.");
            }
            if (string.IsNullOrEmpty(data))
            {
                return null;
            }
            return textTagRegex.Replace(data, m =>
                                              {
                                                  Match match = tagParseRegex.Match(m.Groups["tag"].Value);
                                                  if (match.Success)
                                                  {
                                                      return GetTagReplacementValue(match, dataSet, depth);
                                                  }
                                                  return m.Value;
                                              });
        }


        internal static string GetTagReplacementValue(Match match, IList<DataElement> dataSet, int depth)
        {
            Logger logger = LogManager.GetLogger("Generator");
            string name = match.Groups["name"].Value;
            Index rowIndex = string.IsNullOrEmpty(match.Groups["row"].Value) ? new Index(0) : new Index(match.Groups["row"].Value);
            Index columnIndex = string.IsNullOrEmpty(match.Groups["column"].Value) ? new Index(0) : new Index(match.Groups["column"].Value);
            string format = match.Groups["format"].Value;
            DataElement targetElement = dataSet.FirstOrDefault(e => e.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (targetElement != null)
            {
                object replacementValue = targetElement.Column(columnIndex)[rowIndex];
                if (replacementValue != null)
                {
                    if (replacementValue is string && ((string)replacementValue).IndexOf('#') >= 0)
                    {
                        replacementValue = ParseAndReplace((string)replacementValue, dataSet, ++depth);
                    }
                    if (!string.IsNullOrEmpty(format))
                    {
                        return string.Format("{0:" + format.TrimStart('{'), replacementValue);
                    }
                    return replacementValue.ToString();
                }
                else
                {
                    logger.Debug("Element {0} contains no data", name);
                    return "#" + match.Value + "#";
                }
            }
            logger.Debug("Unable to find data element {0}", name);
            return "#" + match.Value + "#";
        }
    }
}
