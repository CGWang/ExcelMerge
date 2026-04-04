using System.IO;
using System.Windows;
using ExcelMerge.GUI.Views;
using ExcelMerge.GUI.ViewModels;

namespace ExcelMerge.GUI.Commands
{
    public class MergeCommand : ICommand
    {
        public MergeCommandLineOption Option { get; }

        public MergeCommand(MergeCommandLineOption option)
        {
            Option = option;
        }

        public void Execute()
        {
            var window = new MainWindow();
            var diffView = new DiffView();
            var windowViewModel = new MainWindowViewModel(diffView);
            // Show THEIRS on left (src), MINE on right (dst) — consistent with diff layout
            var diffViewModel = new DiffViewModel(Option.TheirsPath, Option.MinePath, windowViewModel);
            window.DataContext = windowViewModel;
            diffView.DataContext = diffViewModel;

            // Set merge-specific properties on the DiffView
            diffView.SetMergeMode(Option.BasePath, Option.OutputPath);

            window.Title = $"ExcelMerge - Merge: {Path.GetFileName(Option.TheirsPath)} \u2194 {Path.GetFileName(Option.MinePath)}";
            window.Closed += (sender, args) => Application.Current.Shutdown();

            App.Current.MainWindow = window;
            window.Show();
        }

        public void ValidateOption()
        {
            if (Option == null)
                throw new Exceptions.ExcelMergeException(true, "Option is null");

            if (string.IsNullOrEmpty(Option.BasePath) || !File.Exists(Option.BasePath))
                throw new Exceptions.ExcelMergeException(true, $"Base file not found: {Option.BasePath}");

            if (string.IsNullOrEmpty(Option.MinePath) || !File.Exists(Option.MinePath))
                throw new Exceptions.ExcelMergeException(true, $"Mine file not found: {Option.MinePath}");

            if (string.IsNullOrEmpty(Option.TheirsPath) || !File.Exists(Option.TheirsPath))
                throw new Exceptions.ExcelMergeException(true, $"Theirs file not found: {Option.TheirsPath}");

            if (string.IsNullOrEmpty(Option.OutputPath))
                throw new Exceptions.ExcelMergeException(true, "Output path is required.");
        }
    }
}
