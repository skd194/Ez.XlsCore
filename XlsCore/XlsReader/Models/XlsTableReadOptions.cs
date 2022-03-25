using System;

namespace Ez.XlsCore
{
    public class XlsTableReadOptions
    {
        public static XlsTableReadOptions Default => new XlsTableReadOptions();

        public XlsTableReadOptions()
            : this(new CellAddress("A", "1"))
        {
        }

        public XlsTableReadOptions(CellAddress startAddress)
            : this(startAddress, null, null)
        {
        }

        public XlsTableReadOptions(CellAddress startAddress,
            Func<HeaderRowContext, RowContext, bool> rowTerminationCondition)
           : this(startAddress, rowTerminationCondition, null)
        {
        }

        public XlsTableReadOptions(
            CellAddress startAddress,
            Func<HeaderRowContext, RowContext, bool> rowTerminationCondition,
            Func<HeaderRowContext, CellContext, bool> columnTerminationCondition)
        {
            StartAddress = startAddress;
            RowTerminationCondition = rowTerminationCondition;
            ColumnTerminationCondition = columnTerminationCondition;
            HasRowTerminationCondition = rowTerminationCondition != null;
            HasColumnTerminationCondition = columnTerminationCondition != null;
        }

        public CellAddress StartAddress { get; }
        internal bool HasRowTerminationCondition;
        internal bool HasColumnTerminationCondition;
        public Func<HeaderRowContext, RowContext, bool> RowTerminationCondition { get; }
        public Func<HeaderRowContext, CellContext, bool> ColumnTerminationCondition { get; }

    }
}
