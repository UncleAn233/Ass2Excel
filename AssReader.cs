using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Ass2Excel
{
    class AssReader
    {
        public const string Dialogue = "Dialogue";
        public const string Comment = "Comment";

        private string pattern = @"(?<type>Dialogue|Comment): \d,(?<start>\d+:\d+:\d+\.\d+),(\d+:\d+:\d+\.\d+),\S*,(?<speaker>{0}),(\d+,\d+,\d+),.*?,(\{{.*\}})?(?<content>.*)";

        public List<AssLine> AssLines { get; private set; }

        public AssReader(string speakers)
        {
            var sp = speakers.Replace("，", "|");
            pattern = String.Format(pattern, sp);
            AssLines = new List<AssLine>();
        }

        public void Read(string path)
        {
            List<string> result = new List<string>();

            var sr = new StreamReader(path);
            string line;
            while((line = sr.ReadLine()) != null)
            {
                var match = Regex.Match(line, pattern);
                if (!match.Success)
                    continue;
                AssLines.Add(new AssLine(match.Groups["type"].ToString(), match.Groups["speaker"].ToString(), match.Groups["start"].ToString(), match.Groups["content"].ToString()));
            }

            sr.Close();
        }

        public int WriteExcel(string path, string fileName)
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Create(path+"\\"+fileName, SpreadsheetDocumentType.Workbook))
            { 
                // 工作表架构
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // 添加工作表集合
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                // 添加工作表到工作簿
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Sheet1"
                };
                sheets.Append(sheet);

                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                sheetData.Append(new Row());
                Row headRow = new Row();
                headRow.Append(
                    CreateCell("A2", "序号", CellValues.String),
                    CreateCell("B2", "时间", CellValues.String),
                    CreateCell("C2", "说话人", CellValues.String),
                    CreateCell("D2", "日文", CellValues.String)
                    );
                sheetData.Append(headRow);

                int count = 0;
                for (var i = 0; i < AssLines.Count; i++)
                {
                    Row row = new Row();
                    if (AssLines[i].Type == Dialogue)
                    {
                        row.Append(
                            //CreateCell(String.Format("A{0}", i + 3), i + 1, CellValues.Number),
                            CreateCell(String.Format("B{0}", i + 3), AssLines[i].Start, CellValues.String),
                            CreateCell(String.Format("C{0}", i + 3), AssLines[i].Speaker, CellValues.String),
                            CreateCell(String.Format("D{0}", i + 3), AssLines[i].Content, CellValues.String)
                            );
                    }

                    else if (AssLines[i].Type == Comment)
                    {
                        count++;
                        row.Append(
                        CreateCell(String.Format("A{0}", i + 3), count, CellValues.Number),
                        //CreateCell(String.Format("B{0}", i + 3), AssLines[i].Start, CellValues.String),
                        CreateCell(String.Format("C{0}", i + 3), AssLines[i].Speaker, CellValues.String),
                        CreateCell(String.Format("D{0}", i + 3), AssLines[i].Content, CellValues.String)
                            );
                    }
                    sheetData.Append(row);
                    Console.WriteLine("写入 " + AssLines[i].ToString());
                }

                // 保存各个部分
                workbookPart.Workbook.Save();
                worksheetPart.Worksheet.Save();

                return AssLines.Count;
            }
        }

        /// <summary>
        /// 提取说话人，省得一个个打了
        /// </summary>
        /// <returns></returns>
        public static string GetSpeakersString(string path)
        {
            HashSet<string> set = new HashSet<string>();
            //string speakerPatther = @"(Dialogue|Comment): \d,(\d+:\d+:\d+\.\d+),(\d+:\d+:\d+\.\d+),\S*,(?<speaker>[^,]*?)(?:\d+,\d+,\d+),.*?,(\{{.*\}})?(.*)";
            string speakerPattern = @"^(?<type>Dialogue|Comment):\s*\d+,(?:\d+:){2}\d+\.\d+,(?:\d+:){2}\d+\.\d+,[^,]*?,(?<speaker>[^,]*?),(?:\d+,\d+,\d+,,.*?(?:{.*?})?)?$";
            var sr = new StreamReader(path);
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                var match = Regex.Match(line, speakerPattern);
                if (!match.Success)
                    continue;
                set.Add(match.Groups["speaker"].ToString());
            }
            Console.Write(set.Count);
            set.Remove("");
            string str = null;
            foreach (var item in set)
            {
                str += item + "，";
            }
            return str.Remove(str.LastIndexOf("，"));
        }

        private static Cell CreateCell(string cellReference, object value, CellValues dataType)
        {
            return new Cell()
            {
                CellReference = cellReference,
                DataType = dataType,
                CellValue = new CellValue(value.ToString())
            };
        }
    }

    public struct AssLine
    {
        public string Type;
        public string Speaker;
        public string Start;
        public string Content;

        public AssLine(string type, string speaker, string start, string content)
        {
            Type = type;
            Speaker = speaker;
            Start = start;
            Content = content;
        }

        public override string ToString()
        {
            return String.Format("说话人：{0}   起始时间：{1}   内容:{2}", Speaker, Start, Content);
        }
    }
}
