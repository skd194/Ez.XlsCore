using System;
using Ez.XlsCore;
using System.Linq;

namespace XlsCore.Demo
{
    class Program
    {
        static void Main(string[] args)
        {
            var file1 = @"C:\Users\SkS\source\repos\ExcelSample\ExcelSample\files\sample.xlsx";

            var file2 = @"C:\Users\SkS\source\repos\ExcelSample\ExcelSample\files\excelwritefile.xlsx";

            var readOptions1 = new XlsTableReadOptions(
                new CellAddress("A", "1"),
                (headerRow, bodyRow) => bodyRow.TryGetCellContext("C", out var cellContext) && cellContext.Value == "Ben100");

            var readOptions2 = new XlsTableReadOptions(
                new CellAddress("A", "1"),
                (headerRow, bodyRow) => bodyRow.IsEmpty,
                null);

            var readOptions3 = new XlsTableReadOptions(new CellAddress("B", "2"));

            using var reader = new XlsReader(file1)
            {
                HeaderRowAction = x =>
                {
                    Console.WriteLine($"Header: {x.Count} " + string.Join(",",
                        x.Cells.Select(context => $"{context.ColumnReference}|{context.Value}")));
                },
                BodyRowAction = x =>
                {
                    Console.WriteLine($"Body: {x.Count} " + string.Join(",",
                        x.Cells.Select(context => $"{context.ColumnReference}|{context.Value}")));
                }
            };

            var result = reader.ReadTable("Second Sheet with name", readOptions2);

            Console.WriteLine();
            Console.WriteLine(result.BodyRowCount);
        }
    }
}
