using System.Collections.Generic;

namespace Ez.XlsCore
{
    public class HeaderRowContext : RowContext
    {
        public HeaderRowContext(
            string rowIndex,
            bool isEmpty,
            IReadOnlyCollection<CellContext> cells)
            : base(rowIndex, isEmpty, cells)
        {
        }
    }
}
