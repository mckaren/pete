using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using McKinsey.PowerPointGenerator.Elements;
using Microsoft.Practices.Unity;

namespace McKinsey.PowerPointGenerator.Commands
{
    public static class CommandManager
    {
        private static IUnityContainer container = new UnityContainer();
        private static Regex regex = new Regex(@"(?<name>[\w_]+)(?:\((?<arguments>.*?)\))?(?:\s+)?", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static List<string> KnownCommandsWithoutArguments = new List<string> { NoContentCommand.Name, FixedCommand.Name, RowHeaderCommand.Name, ColumnHeaderCommand.Name, TransposeCommand.Name, WaterfallCommand.Name };

        static CommandManager()
        {
            container.RegisterType<Command, FormatCommand>(FormatCommand.Name);
            container.RegisterType<Command, VisibleCommand>(VisibleCommand.Name);
            container.RegisterType<Command, LegendCommand>(LegendCommand.Name);
            container.RegisterType<Command, TakeCommand>(TakeCommand.Name);
            container.RegisterType<Command, SkipCommand>(SkipCommand.Name);
            container.RegisterType<Command, SortCommand>(SortCommand.Name);
            container.RegisterType<Command, NoContentCommand>(NoContentCommand.Name);
            container.RegisterType<Command, TransposeCommand>(TransposeCommand.Name);
            container.RegisterType<Command, FormulaCommand>(FormulaCommand.Name);
            container.RegisterType<Command, PageCommand>(PageCommand.Name);
            container.RegisterType<Command, FixedCommand>(FixedCommand.Name);
            container.RegisterType<Command, ReplaceCommand>(ReplaceCommand.Name);
            container.RegisterType<Command, RowHeaderCommand>(RowHeaderCommand.Name);
            container.RegisterType<Command, ColumnHeaderCommand>(ColumnHeaderCommand.Name);
            container.RegisterType<Command, WaterfallCommand>(WaterfallCommand.Name);
            container.RegisterType<Command, ErrorBarCommand>(ErrorBarCommand.Name);
            container.RegisterType<Command, YCommand>(YCommand.Name);
        }


        public static IEnumerable<Command> DiscoverCommands(ShapeElementBase element)
        {
            List<Command> commands = new List<Command>();
            if (string.IsNullOrEmpty(element.CommandString))
            {
                return commands;
            }
            int counter = 0;
            if (regex.Match(element.CommandString).Success)
            {
                foreach (Match match in regex.Matches(element.CommandString))
                {
                    string name = match.Groups["name"].Value.ToUpper();
                    if (container.IsRegistered<Command>(name))
                    {
                        string arguments = match.Groups["arguments"].Value;
                        Command command = container.Resolve<Command>(name);
                        command.ArgumentsString = arguments.Trim();
                        command.TargetElement = element;
                        command.ParseArguments();
                        commands.Add(command);
                    }
                }
            }
            return commands;
        }
    }
}
