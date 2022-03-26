using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Ez.XlsCore
{
    public class XlsReader : IDisposable
    {
        private readonly string[] _sharedStrings;
        private readonly WorkbookPart _workbookPart;
        private readonly SpreadsheetDocument _spreadsheetDocument;

        private HeaderRowContext _headerRowContext;
        private XlsTableReadOptions _xlsTableReadOptions;

        public IReadOnlyCollection<SheetContext> Sheets { get; }
        public Action<RowContext> BodyRowAction { get; set; }
        public Action<HeaderRowContext> HeaderRowAction { get; set; }

        public XlsReader(string path)
        {
            _spreadsheetDocument = SpreadsheetDocument.Open(path, false);
            _workbookPart = _spreadsheetDocument.WorkbookPart;
            _sharedStrings = GetSharedStrings(_workbookPart);
            Sheets = GetSheets(_workbookPart);
        }

        public TableResult ReadTable(int sheetNumber, XlsTableReadOptions options)
        {
            var sheet = Sheets.SingleOrDefault(x => x.Number == sheetNumber);
            return sheet == null
                ? throw new InvalidOperationException("Invalid sheet number")
                : ReadTableById(sheet.Id, options);
        }

        public TableResult ReadTable(string sheetName, XlsTableReadOptions options)
        {
            var sheet = Sheets.SingleOrDefault(x => x.Name == sheetName);
            return sheet == null
                ? throw new InvalidOperationException("Invalid sheet name")
                : ReadTableById(sheet.Id, options);
        }

        public TableResult ReadTable(SheetContext sheet, XlsTableReadOptions options)
        {
            return ReadTableById(sheet.Id, options);
        }

        private TableResult ReadTableById(string sheetId, XlsTableReadOptions options)
        {
            var worksheetPart = _workbookPart.GetPartById(sheetId);
            var reader = OpenXmlReader.Create(worksheetPart);
            _xlsTableReadOptions = options ?? XlsTableReadOptions.Default;
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
                        HeaderRowAction?.Invoke(_headerRowContext);
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
                    if (_xlsTableReadOptions.HasRowTerminationCondition &&
                        _xlsTableReadOptions.RowTerminationCondition(_headerRowContext, rowContext))
                    {
                        if (!rowContext.IsEmpty)
                        {
                            bodyRowCount++;
                            BodyRowAction?.Invoke(rowContext);
                        }

                        break;
                    }

                    bodyRowCount++;
                    BodyRowAction?.Invoke(rowContext);
                }
            }

            return new TableResult(bodyRowCount);
        }
        private static string[] GetSharedStrings(WorkbookPart workbookPart)
        {
            return workbookPart.SharedStringTablePart
                       ?.SharedStringTable
                       .Elements<SharedStringItem>()
                       .Select(x => x.Text.Text)
                       .ToArray()
                   ?? Array.Empty<string>();
        }
        private static SheetContext[] GetSheets(WorkbookPart workbookPart)
        {
            return workbookPart.Workbook.Sheets
                .Cast<Sheet>()
                .Select((sheet, index) => new SheetContext(sheet.Id, index + 1, sheet.Name))
                .ToArray();
        }

        private static string GetRowIndex(OpenXmlReader reader)
        {
            var row = reader.Attributes.SingleOrDefault(x => x.LocalName == "r");
            return row == default
                ? throw new InvalidOperationException(
                    "UnorderedExcelReadNotSupported." +
                    "File data is not ordered explicitly. " +
                    "Reader currently only supports excel file with proper row and cell positioning")
                : row.Value;
        }

        private bool IsContentStartRowIndex(string rowIndex) =>
            rowIndex == _xlsTableReadOptions.StartAddress.Row;

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
                    !_xlsTableReadOptions.HasColumnTerminationCondition)
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
                    if (_xlsTableReadOptions.HasColumnTerminationCondition &&
                        _xlsTableReadOptions.ColumnTerminationCondition(_headerRowContext, cellContext))
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

        private bool IsContentStartColumn(int columnIndex) =>
            columnIndex >= GetColumnIndex(_xlsTableReadOptions.StartAddress.Column);

        private static string GetColumnReference(CellType cell) =>
            Regex.Replace(cell.CellReference.Value.ToUpper(), @"[\d]", string.Empty);

        private static int GetColumnIndex(string columnReference)
        {
            if (string.IsNullOrEmpty(columnReference)) return -1;
            var columnNumber = 0;
            var multiplier = 1;
            foreach (var c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += multiplier * c - 64;
                multiplier *= 26;
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
