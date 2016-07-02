using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Commands;
using McKinsey.PowerPointGenerator.Core.Data;
using McKinsey.PowerPointGenerator.Extensions;
using NLog;

namespace McKinsey.PowerPointGenerator.Elements
{
    public class ChildShapeElement : ShapeElementBase
    {
        public override string TypeName
        {
            get
            {
                return string.Empty;
            }
        }

        public override IEnumerable<Command> PreprocessSwitchCommands(IEnumerable<Command> discoveredCommands)
        {
            return null;
        }
    }
    public abstract class ShapeElementBase
    {
        private static Regex nameParseRegex = new Regex(@"^(?<name>[\w_]+)(?<indexes>(?:\[(?<columns>.*?)\])?(?:\[(?<rows>.*?)\]))?(?:(?:\s+)(?<commands>.*))?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static Regex indexParseRegex = new Regex(@"^(?:\[(?<columns>.*?)\])?(?:\[(?<rows>.*?)\])?$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private bool isMainIndexExtracted = false;

        internal DataElementDescriptor DataDescriptor { get; set; }
        internal List<DataElementDescriptor> AdditionalDataDescriptors { get; set; }
        internal string CommandString { get; set; }
        internal DataElement Data { get; set; }
        internal SlideElement Slide { get; set; }

        public abstract string TypeName { get; }
        public string Name { get; set; }
        public string FullName { get; set; }
        public bool IsReplaced { get; set; }
        public bool UseRowHeaders { get; set; }
        public bool UseColumnHeaders { get; set; }
        public List<Index> RowIndexes { get; private set; }
        public List<Index> ColumnIndexes { get; private set; }
        public List<Command> Commands { get; set; }
        public List<ChildShapeElement> ChildShapes { get; set; }

        public ShapeElementBase()
        {
            ColumnIndexes = new List<Index>();
            RowIndexes = new List<Index>();
            Commands = new List<Command>();
            ChildShapes = new List<ChildShapeElement>();
            UseRowHeaders = false;
            UseColumnHeaders = false;
            DataDescriptor = new DataElementDescriptor();
            AdditionalDataDescriptors = new List<DataElementDescriptor>();
        }

        public bool Parse(string name, SlideElement slide)
        {
            Slide = slide;
            FullName = name.Trim();
            if (!ParseNameString())
            {
                return false;
            }
            if (AdditionalDataDescriptors.Count == 0)
            {
                Name = FullName;
            }
            else
            {
                Name = AdditionalDataDescriptors[0].Name;
                foreach (var additionalDescriptor in AdditionalDataDescriptors)
                {
                    ExtractIndexes(additionalDescriptor);
                }
            }
            return true;
        }

        public bool ParseNameString()
        {
            string nameToParse = FullName.Replace('“', '"').Replace('”', '"');
            DataElementDescriptor dataDescriptor = DataDescriptor;
            if (ParseIntoDescriptor(nameToParse))
            {
                //DataElementDescriptor additionalDescriptor;
                //while (ParseIntoDescriptor(CommandString, additionalDescriptor = new DataElementDescriptor()))
                //{
                //    AdditionalDataDescriptors.Add(additionalDescriptor);
                //}
                return true;
            }
            return false;
        }

        private bool ParseIntoDescriptor(string nameToParse)
        {
            Match match = nameParseRegex.Match(nameToParse);
            if (match.Success)
            {
                string indexes = match.Groups["indexes"].Value;
                CommandString = match.Groups["commands"].Value;
                foreach (var fragment in indexes.Split(new char[] { ':' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    string name = match.Groups["name"].Value;
                    Match indexMatch = indexParseRegex.Match(fragment);
                    if (!CommandManager.KnownCommandsWithoutArguments.Contains(name.ToUpper()) && indexMatch.Success)
                    {
                        DataElementDescriptor dataDescriptor = new DataElementDescriptor();
                        dataDescriptor.Name = name;
                        dataDescriptor.RowIndexesString = indexMatch.Groups["rows"].Value;
                        dataDescriptor.ColumnIndexesString = indexMatch.Groups["columns"].Value;
                        AdditionalDataDescriptors.Add(dataDescriptor);
                    }           
                }
                return true;
            }
            return false;
        }
        public List<T> CommandsOf<T>() where T : Command
        {
            return Commands.Where(c => (c as T) != null).Cast<T>().ToList();
        }

        protected void ExtractIndexes(DataElementDescriptor dataDescriptor)
        {
            if (!isMainIndexExtracted)
            {
                dataDescriptor.ExtractColumnIndexes(ColumnIndexes);
                dataDescriptor.ExtractRowIndexes(RowIndexes);
                isMainIndexExtracted = true;
            } else
            {
                ChildShapeElement childShape = new ChildShapeElement();
                dataDescriptor.ExtractColumnIndexes(childShape.ColumnIndexes);
                dataDescriptor.ExtractRowIndexes(childShape.RowIndexes);
                ChildShapes.Add(childShape);
            }
        }

        /// <summary>
        /// Calls the command manager to discover commands and then asks inherited classes to preprocess them. Some elements may want to use certain commands as switches.
        /// </summary>
        public virtual void DiscoverCommands()
        {
            if (!string.IsNullOrEmpty(CommandString))
            {
                var discoveredCommands = CommandManager.DiscoverCommands(this);
                Commands.Clear();
                discoveredCommands = PreprocessSwitchCommandsBase(discoveredCommands);
                discoveredCommands = PreprocessSwitchCommands(discoveredCommands);
                Commands.AddRange(discoveredCommands);
            }
            CheckCommandsForIndexes();
        }

        internal void CheckCommandsForIndexes()
        {
            bool isTransposed = false;
            foreach (Command command in Commands)
            {
                if ((command as TransposeCommand) != null)
                {
                    isTransposed = !isTransposed;
                }
                var indexedCommand = command as IUseIndexes;
                if (indexedCommand != null)
                {
                    foreach (Index usedIndex in indexedCommand.UsedIndexes)
                    {
                        if (isTransposed)
                        {
                            if (!RowIndexes.Any(i => i == usedIndex || i.IsAll))
                            {
                                usedIndex.IsHidden = true;
                                RowIndexes.Add(usedIndex);
                            }
                        }
                        else
                        {
                            if (!ColumnIndexes.Any(i => i == usedIndex || i.IsAll))
                            {
                                usedIndex.IsHidden = true;
                                ColumnIndexes.Add(usedIndex);
                            }
                        }
                    }
                }
                var errorBarCommand = command as ErrorBarCommand;
                if (errorBarCommand != null)
                {
                    ColumnIndexes.Add(errorBarCommand.MinusIndex);
                    if (errorBarCommand.PlusIndex != null)
                    {
                        ColumnIndexes.Add(errorBarCommand.PlusIndex);
                    }
                }
                var yCommand = command as YCommand;
                if (yCommand != null)
                {
                    if (isTransposed)
                    {
                        if (!RowIndexes.Any(i => i == yCommand.Index || i.IsAll))
                        {
                            RowIndexes.Add(yCommand.Index);
                        }
                    }
                    else
                    {
                        ColumnIndexes.Add(yCommand.Index);
                    }
                }
            }
        }

        public virtual void FindShapeData(IList<DataElement> data)
        {
            if (Data != null)
            {
                return;
            }
            DataElement applicableData = data.FirstOrDefault(d => d.Name.Equals(Name, StringComparison.OrdinalIgnoreCase));
            if (applicableData != null)
            {
                Data = applicableData.Clone();
                foreach (DataElementDescriptor additionalDataDescriptor in AdditionalDataDescriptors)
                {
                    DataElement additionalDataElement = data.FirstOrDefault(d => d.Name.Equals(additionalDataDescriptor.Name, StringComparison.OrdinalIgnoreCase)); ;
                    if (additionalDataElement != null && !additionalDataElement.Name.Equals(Data.Name, StringComparison.OrdinalIgnoreCase))
                    {
                        Data.MergeWith(additionalDataElement);
                    }
                }
            }
            else
            {
                Logger logger = LogManager.GetLogger("Generator");
                logger.Debug("Unable to find data for {0} on slide {1}.", FullName, Slide.Number);
            }
        }

        internal virtual IEnumerable<Command> PreprocessSwitchCommandsBase(IEnumerable<Command> discoveredCommands)
        {
            var processedCommands = new List<Command>(discoveredCommands);
            if (processedCommands.Any(c => (c as RowHeaderCommand) != null))
            {
                UseRowHeaders = true;
                processedCommands.Remove(processedCommands.First(c => (c as RowHeaderCommand) != null));
            }
            if (processedCommands.Any(c => (c as ColumnHeaderCommand) != null))
            {
                UseColumnHeaders = true;
                processedCommands.Remove(processedCommands.First(c => (c as ColumnHeaderCommand) != null));
            }
            return processedCommands;
        }

        internal virtual void ProcessCommands(DataElement data)
        {
            foreach (Command command in Commands)
            {
                command.ApplyToData(data);
            }
        }

        public abstract IEnumerable<Command> PreprocessSwitchCommands(IEnumerable<Command> discoveredCommands);
    }
}