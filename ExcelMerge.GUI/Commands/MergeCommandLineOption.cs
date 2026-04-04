using System.IO;
using CommandLine;

namespace ExcelMerge.GUI.Commands
{
    [Verb("merge", HelpText = "Three-way merge of Excel files.")]
    public class MergeCommandLineOption
    {
        [Option("base", Required = true, HelpText = "Base (common ancestor) file path.")]
        public string BasePath { get; set; } = string.Empty;

        [Option("mine", Required = true, HelpText = "Mine (local) file path.")]
        public string MinePath { get; set; } = string.Empty;

        [Option("theirs", Required = true, HelpText = "Theirs (remote) file path.")]
        public string TheirsPath { get; set; } = string.Empty;

        [Option("output", Required = true, HelpText = "Output merged file path.")]
        public string OutputPath { get; set; } = string.Empty;

        [Option("auto", HelpText = "Auto-merge non-conflicting changes. Open GUI only if conflicts exist.")]
        public bool Auto { get; set; }

        [Value(0, MetaName = "base", HelpText = "Base file path (positional).")]
        public string PositionalBase { get; set; }

        [Value(1, MetaName = "mine", HelpText = "Mine file path (positional).")]
        public string PositionalMine { get; set; }

        [Value(2, MetaName = "theirs", HelpText = "Theirs file path (positional).")]
        public string PositionalTheirs { get; set; }

        [Value(3, MetaName = "output", HelpText = "Output file path (positional).")]
        public string PositionalOutput { get; set; }

        public void ConvertToFullPath()
        {
            if (string.IsNullOrEmpty(BasePath) && !string.IsNullOrEmpty(PositionalBase))
                BasePath = PositionalBase;
            if (string.IsNullOrEmpty(MinePath) && !string.IsNullOrEmpty(PositionalMine))
                MinePath = PositionalMine;
            if (string.IsNullOrEmpty(TheirsPath) && !string.IsNullOrEmpty(PositionalTheirs))
                TheirsPath = PositionalTheirs;
            if (string.IsNullOrEmpty(OutputPath) && !string.IsNullOrEmpty(PositionalOutput))
                OutputPath = PositionalOutput;

            BasePath = !string.IsNullOrEmpty(BasePath) ? Path.GetFullPath(BasePath) : BasePath;
            MinePath = !string.IsNullOrEmpty(MinePath) ? Path.GetFullPath(MinePath) : MinePath;
            TheirsPath = !string.IsNullOrEmpty(TheirsPath) ? Path.GetFullPath(TheirsPath) : TheirsPath;
            OutputPath = !string.IsNullOrEmpty(OutputPath) ? Path.GetFullPath(OutputPath) : OutputPath;
        }
    }
}
