using System.Collections.Generic;
using System.Linq;

namespace Ez.XlsCore
{
    public class RowContext
    {
        private readonly Dictionary<string, CellContext> _cells;
        public RowContext(
            string rowIndex,
             bool isRowEmpty,
            IEnumerable<CellContext> cells)
        {
            RowIndex = rowIndex;
            _cells = cells.ToDictionary(x => x.ColumnReference);
            Count = _cells.Count;
            IsEmpty = isRowEmpty;
        }

        public string RowIndex { get; }
        public int Count { get; }
        public bool IsEmpty { get; }

        public IReadOnlyCollection<CellContext> Cells => _cells.Values;
        public bool TryGetCellContext(string columnReference, out CellContext cellContext) =>
            _cells.TryGetValue(columnReference, out cellContext);

    }
}
