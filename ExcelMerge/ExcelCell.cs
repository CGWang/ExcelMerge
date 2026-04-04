namespace ExcelMerge
{
    public class ExcelCell
    {
        public string Value { get; private set; }
        public string Formula { get; private set; }
        public string Comment { get; private set; }
        public int OriginalColumnIndex { get; private set; }
        public int OriginalRowIndex { get; private set; }

        public ExcelCell(string value, int originalColumnIndex, int originalRowIndex,
                         string formula = null, string comment = null)
        {
            Value = value;
            Formula = formula ?? string.Empty;
            Comment = comment ?? string.Empty;
            OriginalColumnIndex = originalColumnIndex;
            OriginalRowIndex = originalRowIndex;
        }
    }
}
