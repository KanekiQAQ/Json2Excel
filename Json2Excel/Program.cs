using OfficeOpenXml;
using System;
using System.IO;
using System.Text;
using System.Text.Json;

namespace Json2Excel
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("请输入文件名：");
                var fileName = Console.ReadLine();
                var content = GetContent(fileName);
                var res = Convert2Excel(fileName, content);
                ExportXlsx(fileName, res);
                Console.WriteLine($"生成{fileName}.xlsx完成");
                Console.ReadKey();
            }
        }

        private static string GetContent(string fileName)
        {
            string content = "";
            var filePath = Environment.CurrentDirectory + $"\\{fileName}.json";
            if (File.Exists(filePath))
            {
                content = File.ReadAllText(filePath);
                byte[] bytes = Encoding.UTF8.GetBytes(content);
                content = Encoding.UTF8.GetString(bytes);
            }
            return content;
        }

        private static void ExportXlsx(string fileName, byte[] content)
        {
            string filePath = Environment.CurrentDirectory + $"\\{fileName}.xlsx";
            FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
            fileStream.Write(content, 0, content.Length);
            fileStream.Flush();
            fileStream.Close();
        }

        private static byte[] Convert2Excel(string fileName, string content)
        {
            var jsonDocument = JsonDocument.Parse(content, new JsonDocumentOptions
            {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip
            });
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(fileName);
                int row = 1;
                foreach (var element in jsonDocument.RootElement.EnumerateArray())
                {
                    int col = 1;
                    foreach (var obj in element.EnumerateObject())
                    {
                        if (row == 1)
                        {
                            worksheet.Cells[1, col].Value = obj.Name;
                            worksheet.Cells[1, col].Style.Font.Bold = true;
                            worksheet.Cells[1, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheet.Cells[1, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[1, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(211, 211, 211));
                            worksheet.Cells[2, col].Value = obj.Value.ToString();
                            col++;
                        }
                        else
                        {
                            worksheet.Cells[row, col].Value = obj.Value.ToString();
                            col++;
                        }
                    }
                    if (row == 1) 
                    {
                        row = 2;
                    }
                    row++;                    
                }
                worksheet.Cells.AutoFitColumns();
                package.Save();
                return package.GetAsByteArray();
            }
        }
    }
}
