namespace Ez.XlsCore
{
    public class CellAddress
    {
        public CellAddress(string column, string row)
        {
            Row = row;
            Column = column;
        }

        public string Column { get; }
        public string Row { get; }

        public static implicit operator string(CellAddress address) => address.ToString();
        public override string ToString() => $"{Column}{Row}";
    }
}
