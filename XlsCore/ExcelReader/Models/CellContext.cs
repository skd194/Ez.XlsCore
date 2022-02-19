namespace Ez.XlsCore
{
    public class CellContext
    {
        public CellContext(string value, string columnReference, bool isEmpty)
        {
            Value = value;
            ColumnReference = columnReference;
            IsEmpty = isEmpty;
        }

        public string Value { get; }
        public string ColumnReference { get; }
        public bool IsEmpty { get; }
    }
}
