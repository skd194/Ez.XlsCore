using System;

namespace Ez.XlsCore
{
    public class ReadOptions
    {
        public ReadOptions(CellAddress startAddress)
            : this(startAddress,
                  (_, bodyRow) => true,
                  (_, cell) => true)
        {
        }

        public ReadOptions(CellAddress startAddress,
            Func<HeaderRowContext, RowContext, bool> rowTerminationCondition)
           : this(startAddress,
             rowTerminationCondition,
             (_, cell) => true)
        {
        }

        public ReadOptions(
            CellAddress startAddress,
            Func<HeaderRowContext, RowContext, bool> rowTerminationCondition,
            Func<HeaderRowContext, CellContext, bool> columnTerminationCondition)
        {
            StartAddress = startAddress;
            RowTerminationCondition = rowTerminationCondition;
            ColumnTerminationCondition = columnTerminationCondition;
        }

        public CellAddress StartAddress { get; }
        public Func<HeaderRowContext, RowContext, bool> RowTerminationCondition { get; }
        public Func<HeaderRowContext, CellContext, bool> ColumnTerminationCondition { get; }

    }
}
