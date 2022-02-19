using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Ez.XlsCore
{
    public partial class XlsReader : IDisposable
    {
        private readonly SpreadsheetDocument _spreadsheetDocument;

        private readonly WorkbookPart _workbookPart;

        private readonly WorksheetPart _worksheetPart;

        private readonly ReadOptions _readOptions;

        private readonly string[] _sharedStrings;

        private HeaderRowContext _headerRowContext;

        public XlsReader(string path, ReadOptions options)
        {
            _spreadsheetDocument = SpreadsheetDocument.Open(path, false);
            _workbookPart = _spreadsheetDocument.WorkbookPart;
            _worksheetPart = _workbookPart.WorksheetParts.First();
            _sharedStrings = _workbookPart.SharedStringTablePart.SharedStringTable
                .Elements<SharedStringItem>()
                .Select(x => x.Text.Text)
                .ToArray();
            _readOptions = options ?? throw new ArgumentNullException(nameof(options));
        }

        private bool IsContentStartRowIndex(string rowIndex) => rowIndex == _readOptions.StartAddress.Row;

        public TableResult ReadTable(
            string sheetName,
            Action<HeaderRowContext> headerRowAction,
            Action<RowContext> bodyRowAction)
        {
            var reader = OpenXmlReader.Create(_worksheetPart);
            var skipRow = true;
            var bodyRowCount = 0;
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Row)) continue;
                var rowIndex = GetRowIndex(reader);
                reader.ReadFirstChild();
                if (skipRow)
                {
                    if (IsContentStartRowIndex(rowIndex))
                    {
                        _headerRowContext = ReadHeaderRow(reader, rowIndex);
                        headerRowAction(_headerRowContext);
                        skipRow = false;
                    }
                    else
                    {
                        SkipStream(reader);
                    }
                }
                else
                {
                    var result = ReadRow(reader, _headerRowContext.Count);
                    var rowContext = new RowContext(rowIndex, result.IsEmpty, result.Cells);
                    if (_readOptions.HasRowTerminationCondition &&
                        _readOptions.RowTerminationCondition(_headerRowContext, rowContext))
                    {
                        if (!rowContext.IsEmpty)
                        {
                            bodyRowCount++;
                            bodyRowAction(rowContext);
                        }
                        break;
                    }
                    bodyRowCount++;
                    bodyRowAction(rowContext);
                }
            }
            return new TableResult(bodyRowCount);
        }

        private static string GetRowIndex(OpenXmlReader reader) =>
            reader.Attributes.Single(x => x.LocalName == "r").Value;

        private HeaderRowContext ReadHeaderRow(
            OpenXmlReader reader,
            string rowIndex)
        {
            var headerRowResult = ReadRow(reader);
            var headerRowContext = new HeaderRowContext(
                rowIndex,
                headerRowResult.IsEmpty,
                headerRowResult.Cells);
            return headerRowContext;
        }

        private static void SkipStream(OpenXmlReader reader)
        {
            do
            {
            } while (reader.ReadNextSibling());
        }

        private RowResult ReadRow(OpenXmlReader reader, int? rowItemsCount = null)
        {
            var cells = new List<CellContext>();
            var isRowEmpty = true;
            var startCellRead = false;
            var stopCellRead = false;
            var itemCount = 1;
            do
            {
                if (stopCellRead) continue;
                if (rowItemsCount.HasValue &&
                    itemCount > rowItemsCount.Value &&
                    !_readOptions.HasColumnTerminationCondition)
                {
                    continue;
                }
                if (reader.ElementType != typeof(Cell)) continue;
                var cell = (Cell)reader.LoadCurrentElement();
                var columnReference = GetColumnReference(cell);
                var columnIndex = GetColumnIndex(columnReference);
                if (!startCellRead && IsContentStartColumn(columnIndex)) startCellRead = true;
                if (startCellRead)
                {
                    var value = GetCellRawValue(cell);
                    var cellContext = new CellContext(
                        value,
                        columnReference,
                        string.IsNullOrEmpty(value),
                        columnIndex);
                    if (_readOptions.HasColumnTerminationCondition &&
                        _readOptions.ColumnTerminationCondition(_headerRowContext, cellContext))
                    {
                        stopCellRead = true;
                        continue;
                    }
                    cells.Add(cellContext);
                    if (!cellContext.IsEmpty) isRowEmpty = false;
                    itemCount++;
                }
            } while (reader.ReadNextSibling());

            return new RowResult(cells, isRowEmpty);
        }

        private string GetCellRawValue(CellType cell) =>
            cell.DataType != null && cell.DataType == CellValues.SharedString
                ? _sharedStrings[int.Parse(cell.CellValue.InnerText)]
                : cell.CellValue?.InnerText;

        private bool IsContentStartColumn(int columnIndex)
        {
            return columnIndex >= GetColumnIndex(_readOptions.StartAddress.Column);
        }

        private static string GetColumnReference(CellType cell) =>
            Regex.Replace(cell.CellReference.Value.ToUpper(), @"[\d]", string.Empty);

        private static int GetColumnIndex(string columnReference)
        {
            if (string.IsNullOrEmpty(columnReference)) return -1;
            int columnNumber = 0;
            int mulitplier = 1;
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * c - 64;
                mulitplier *= 26;
            }
            return columnNumber;
        }

        public void Dispose()
        {
            _spreadsheetDocument.Dispose();
        }

        private class RowResult
        {
            public RowResult(IReadOnlyCollection<CellContext> cells, bool isEmpty)
            {
                Cells = cells;
                IsEmpty = isEmpty;
            }

            public IReadOnlyCollection<CellContext> Cells { get; }
            public bool IsEmpty { get; }
        }
    }
}
