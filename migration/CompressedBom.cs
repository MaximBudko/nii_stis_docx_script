using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace Migration
{
    public static class CompressedBom
    {
        public static void ProcessExcel(string inputFile, string outputFile)
        {
            using (var package = new ExcelPackage(new FileInfo(inputFile)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var headers = new List<string>();
                var data = new Dictionary<string, (int quantity, List<string> numbers, List<string> fullRow)>();

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Text);
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var key = worksheet.Cells[row, 5].Text;
                    if (string.IsNullOrEmpty(key)) continue;

                    var number = worksheet.Cells[row, 1].Text;
                    var quantity = int.Parse(worksheet.Cells[row, 4].Text);
                    var fullRow = new List<string>();

                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        fullRow.Add(worksheet.Cells[row, col].Text);
                    }

                    if (data.ContainsKey(key))
                    {
                        data[key] = (data[key].quantity + quantity, data[key].numbers, data[key].fullRow);
                        data[key].numbers.Add(number);
                    }
                    else
                    {
                        data[key] = (quantity, new List<string> { number }, fullRow);
                    }
                }

                using (var outputPackage = new ExcelPackage(new FileInfo(outputFile)))
                {
                    var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Sheet1");

                    for (int col = 0; col < headers.Count; col++)
                    {
                        outputWorksheet.Cells[1, col + 1].Value = headers[col];
                    }

                    int outputRow = 2;
                    foreach (var entry in data)
                    {
                        var newRow = entry.Value.fullRow.ToArray();
                        newRow[0] = string.Join(", ", entry.Value.numbers);
                        newRow[3] = entry.Value.quantity.ToString();

                        for (int col = 0; col < newRow.Length; col++)
                        {
                            outputWorksheet.Cells[outputRow, col + 1].Value = newRow[col];
                        }

                        outputRow++;
                    }

                    outputPackage.Save();
                }
            }
        }
    }
}
