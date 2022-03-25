using Ez.XlsCore;
using NUnit.Framework;
using System.Collections.Generic;
using System.Linq;

namespace XlsCore.Tests
{
    public class XlsReadTableSimpleTests
    {
        private string _fileName;
        private string[] _sheetColumnsWithContent;
        private string _sheetHeaderRow;
        private string[] _sheetBodyRows;
        private string[] _sheetHeaderCellValues;
        private int _sheetBodyRowCount;
        private int[] _sheetColumnIndexes;
        private string[] _sheetDecimalColumnValues;

        [SetUp]
        public void Setup()
        {
            _fileName = @"..\..\..\Reader\XlsFiles\SimpleTable.xlsx";
            _sheetColumnsWithContent = new[] { "A", "B", "C", "D", "E" };
            _sheetHeaderRow = "1";
            _sheetBodyRows = new[] { "2", "3", "4", "5", "6", "7" };
            _sheetHeaderCellValues = new[] { "EmployeeCode", "EmployeeName", "Salary", "DateOfJoining", "NoOfLeaves" };
            _sheetBodyRowCount = 6;
            _sheetColumnIndexes = new[] { 1, 2, 3, 4, 5 };
            _sheetDecimalColumnValues = new[] {"1500.5", "2400.75", "2000.67", "480.0", "1000.99", "1767.987"};
        }

        [Test]
        public void ReadTable_WhenCalled_WithNoTerminationConditions_CheckHeaderRowIndex()
        {
            var readOptions = new XlsTableReadOptions();
            using var reader = new XlsReader(_fileName, readOptions);
            string headerRow = null;
            reader.ReadTable(
                header => { headerRow = header.RowIndex; },
                body => { });
            Assert.AreEqual(headerRow, _sheetHeaderRow);
        }

        [Test]
        public void ReadTable_WhenCalled_WithNoTerminationConditions_CheckHeaderCellValues()
        {
            var readOptions = new XlsTableReadOptions();
            using var reader = new XlsReader(_fileName, readOptions);
            string[] headerCellValues = null;
            reader.ReadTable(
                header => headerCellValues = header.Cells.Select(x => x.Value).ToArray(),
                body => { });

            CollectionAssert.AreEqual(headerCellValues, _sheetHeaderCellValues);
        }

        [Test]
        public void ReadTable_WhenCalled_WithNoTerminationConditions_CheckBodyRows()
        {
            var readOptions = new XlsTableReadOptions();

            using var reader = new XlsReader(_fileName, readOptions);

            var bodyRowIndexes = new List<string>();

            reader.ReadTable(
                _ => { },
                body => bodyRowIndexes.Add(body.RowIndex));

            CollectionAssert.AreEqual(bodyRowIndexes, _sheetBodyRows);
        }

        [Test]
        public void ReadTable_WhenCalled_WithNoTerminationConditions_CheckBodyColumns()
        {
            var readOptions = new XlsTableReadOptions();

            using var reader = new XlsReader(_fileName, readOptions);

            var bodyColumns = new List<string>();

            reader.ReadTable(
                _ => { },
                body => bodyColumns = body.Cells.Select(x => x.ColumnReference).ToList());

            CollectionAssert.AreEqual(bodyColumns, _sheetColumnsWithContent);
        }

        [Test]
        public void ReadTable_WhenCalled_WithNoTerminationConditions_ReturnBodyRowCount()
        {
            var readOptions = new XlsTableReadOptions();

            using var reader = new XlsReader(_fileName, readOptions);

            var result = reader.ReadTable(_ => { }, _ => { });

            Assert.AreEqual(result.BodyRowCount, _sheetBodyRowCount);
        }

        [Test]
        public void ReadTable_WhenCalled_WithNoTerminationConditions_CheckColumnIndex()
        {
            var readOptions = new XlsTableReadOptions();

            using var reader = new XlsReader(_fileName, readOptions);

            var columnIndexes = new List<int>();

            reader.ReadTable(_ => { },
                body => columnIndexes = body.Cells.Select(x => x.ColumnIndex).ToList());

            Assert.AreEqual(columnIndexes, _sheetColumnIndexes);
        }

        [Test]
        public void ReadTable_WhenCalled_WithNoTerminationConditions_CheckBodyDecimalColumnValue()
        {
            var readOptions = new XlsTableReadOptions();
            using var reader = new XlsReader(_fileName, readOptions);
            var bodyColumnValues = new List<string>();
            reader.ReadTable(
                _ => { },
                body => bodyColumnValues.Add(body.Cells.First(x => x.ColumnReference == "C").Value));

            CollectionAssert.AreEqual(bodyColumnValues, _sheetDecimalColumnValues);
        }
    }
}