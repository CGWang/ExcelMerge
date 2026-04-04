namespace ExcelMerge.GUI.Commands
{
    public interface ICommand
    {
        void Execute();
        void ValidateOption();
    }
}
