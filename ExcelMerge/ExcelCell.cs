namespace ExcelMerge
{
    public class ExcelCell
    {
        public string Value { get; private set; }
        public string Formula { get; private set; }
        public int OriginalColumnIndex { get; private set; }
        public int OriginalRowIndex { get; private set; }

        public ExcelCell(string value, int originalColumnIndex, int originalRowIndex)
        {
            Value = value;
            Formula = string.Empty;
            OriginalColumnIndex = originalColumnIndex;
            OriginalRowIndex = originalRowIndex;
        }

        public ExcelCell(string value, string formula, int originalColumnIndex, int originalRowIndex)
        {
            Value = value;
            Formula = formula ?? string.Empty;
            OriginalColumnIndex = originalColumnIndex;
            OriginalRowIndex = originalRowIndex;
        }
    }
}
