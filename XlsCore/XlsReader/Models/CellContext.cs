namespace Ez.XlsCore
{
    public class CellContext
    {
        public CellContext(string value, string columnReference, bool isEmpty, int columnIndex)
        {
            Value = value;
            ColumnReference = columnReference;
            IsEmpty = isEmpty;
            ColumnIndex = columnIndex;
        }

        public string Value { get; }
        public string ColumnReference { get; }
        public bool IsEmpty { get; }
        public int ColumnIndex { get; }
    }
}
