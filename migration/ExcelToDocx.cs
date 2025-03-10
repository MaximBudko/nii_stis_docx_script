using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using Xceed.Words.NET;

namespace Migration
{
    public static class ExcelToDocx
    {
        public static void GenerateDocx(string templatePath, string excelPath, Dictionary<string, string> fieldValues, string outputPath)
        {
            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var data = new List<string[]>();

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    var rowData = new string[4];
                    for (int col = 1; col <= 4; col++)
                    {
                        rowData[col - 1] = worksheet.Cells[row, col].Text;
                    }
                    data.Add(rowData);
                }

                using (var doc = DocX.Load(templatePath))
                {
                    foreach (var field in fieldValues)
                    {
                        doc.ReplaceText($"{{{{{field.Key}}}}}", field.Value);
                    }

                    var table = doc.Tables[0];
                    foreach (var row in data)
                    {
                        var newRow = table.InsertRow();
                        for (int i = 0; i < row.Length; i++)
                        {
                            newRow.Cells[i].Paragraphs[0].Append(row[i]);
                        }
                    }

                    doc.SaveAs(outputPath);
                }
            }
        }
    }
}
