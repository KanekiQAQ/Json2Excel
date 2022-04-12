using OfficeOpenXml;
using System;
using System.Collections.Generic;
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
            var dictionary = new Dictionary<string, List<string>>();
            var jsonDocument = JsonDocument.Parse(content, new JsonDocumentOptions
            {
                AllowTrailingCommas = true,
                CommentHandling = JsonCommentHandling.Skip
            });
            foreach (var element in jsonDocument.RootElement.EnumerateArray())
            {
                foreach (var obj in element.EnumerateObject())
                {
                    if (!dictionary.ContainsKey(obj.Name))
                    {
                        var list = new List<string>();
                        list.Add(obj.Name);
                        dictionary.Add(obj.Name, list);
                    }
                }
            }
            foreach (var element in jsonDocument.RootElement.EnumerateArray())
            {
                foreach (var dic in dictionary)
                {
                    JsonElement j;
                    if (element.TryGetProperty(dic.Key, out j))
                    {
                        dictionary[dic.Key].Add(j.ToString());
                    }
                    else
                    {
                        dictionary[dic.Key].Add(null);
                    }
                }
            }
            using (ExcelPackage package = new ExcelPackage())
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(fileName);
                int col = 1;
                foreach (var dic in dictionary)
                {
                    int row = 1;
                    for (int i = 0; i < dictionary[dic.Key].Count; i++)
                    {
                        if (row == 1)
                        {
                            worksheet.Cells[row, col].Value = dictionary[dic.Key][i];
                            worksheet.Cells[row, col].Style.Font.Bold = true;
                            worksheet.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            worksheet.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[row, col].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(211, 211, 211));
                        }
                        else
                        {
                            worksheet.Cells[row, col].Value = dictionary[dic.Key][i];
                        }
                        row++;
                    }
                    col++;
                }
                worksheet.Cells.AutoFitColumns();
                package.Save();
                return package.GetAsByteArray();
            }
        }
    }
}
