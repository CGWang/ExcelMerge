using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using CommandLine;

namespace ExcelMerge.GUI.Commands
{
    [Verb("diff", isDefault: true, HelpText = "Compare two files.")]
    public class CommandLineOption
    {
        [Option('s', "src-path", HelpText = "Source file path.")]
        public string SrcPath { get; set; } = string.Empty;

        [Option('d', "dst-path", HelpText = "Dest file path.")]
        public string DstPath { get; set; } = string.Empty;

        [Option('c', "external-cmd", HelpText = "External command for unsupported file types.")]
        public string ExternalCommand { get; set; } = string.Empty;

        [Option('i', "immediately-execute-external-cmd", HelpText = "Execute external cmd without error dialog.")]
        public bool ImmediatelyExecuteExternalCommand { get; set; }

        [Option('w', "wait-external-cmd", HelpText = "Wait for the external process to finish.")]
        public bool WaitExternalCommand { get; set; }

        [Option('v', "validate-extension", HelpText = "Validate extension before open file.")]
        public bool ValidateExtension { get; set; }

        [Option('e', "empty-file-name", HelpText = "Empty file name.")]
        public string EmptyFileName { get; set; } = string.Empty;

        [Option('k', "keep-file-history", HelpText = "Don't add recent files.")]
        public bool KeepFileHistory { get; set; }

        [Option("quit-on-close", HelpText = "Quit application when diff window is closed.")]
        public bool QuitOnClose { get; set; }

        [Option("base-name", HelpText = "Display name for source/base file.")]
        public string BaseName { get; set; } = string.Empty;

        [Option("mine-name", HelpText = "Display name for destination/mine file.")]
        public string MineName { get; set; } = string.Empty;

        [Option("readonly-left", HelpText = "Treat the left (source) file as read-only.")]
        public bool ReadonlyLeft { get; set; }

        [Value(0, MetaName = "source", HelpText = "Source file path (positional).")]
        public string PositionalSrc { get; set; }

        [Value(1, MetaName = "destination", HelpText = "Destination file path (positional).")]
        public string PositionalDst { get; set; }

        public CommandType MainCommand => CommandType.Diff;

        public void ConvertToFullPath()
        {
            if (string.IsNullOrEmpty(SrcPath) && !string.IsNullOrEmpty(PositionalSrc))
                SrcPath = PositionalSrc;

            if (string.IsNullOrEmpty(DstPath) && !string.IsNullOrEmpty(PositionalDst))
                DstPath = PositionalDst;

            SrcPath = !string.IsNullOrEmpty(SrcPath) ? Path.GetFullPath(SrcPath) : SrcPath;
            DstPath = !string.IsNullOrEmpty(DstPath) ? Path.GetFullPath(DstPath) : DstPath;
        }
    }
}
