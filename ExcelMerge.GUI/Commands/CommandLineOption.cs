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

        public CommandType MainCommand => CommandType.Diff;

        public void ConvertToFullPath()
        {
            SrcPath = !string.IsNullOrEmpty(SrcPath) ? Path.GetFullPath(SrcPath) : SrcPath;
            DstPath = !string.IsNullOrEmpty(DstPath) ? Path.GetFullPath(DstPath) : DstPath;
        }
    }
}
