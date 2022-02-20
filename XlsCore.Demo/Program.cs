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
            var file2 = @"C:\Users\SkS\Desktop\MyWork\Ez.XlsCore\XlsCore.Tests\Reader\XlsFiles\SimpleTable.xlsx";


            var readOptions1 = new XlsReadOptions(
                new CellAddress("A", "1"),
                (headerRow, bodyRow) => bodyRow.TryGetCellContext("C", out var cellContext) && cellContext.Value == "Ben100");

            var readOptions2 = new XlsReadOptions(
                new CellAddress("A", "1"),
                (headerRow, bodyRow) => bodyRow.IsEmpty,
                null);

            var readOptions3 = new XlsReadOptions(new CellAddress("B", "2"));

            using var reader = new XlsReader(file2, readOptions2);

            var result = reader.ReadTable(
                x => { Console.WriteLine($"Header: {x.Count} " + string.Join(",", x.Cells.Select(x => $"{x.ColumnReference}|{x.Value}"))); },
                x => { Console.WriteLine($"Body: {x.Count} " + string.Join(",", x.Cells.Select(x => $"{x.ColumnReference}|{x.Value}"))); });

            Console.WriteLine();
            Console.WriteLine(result.BodyRowCount);
        }
    }
}
