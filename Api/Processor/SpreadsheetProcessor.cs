using System;
using App.Api.Dto;
using DocumentFormat.OpenXml.Packaging;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using App.Api.Processor.Storage;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Extensions.Logging.Abstractions;

namespace App.Api.Processor
{
    public class SpreadsheetProcessor: AbstractDocumentProcessor<CellSegmentConfig>
    {
        public override IDocument ExtractDocument(string filePath)
        {
            SpreadsheetDocument Doc = SpreadsheetDocument.Open(filePath, false);
            var sheetList = Doc.WorkbookPart.Workbook.Sheets;
            string content = null;
            Segment<CellSegmentConfig> textSegment = new Segment<CellSegmentConfig>();
            List<Style> styleList = new List<Style>();
            int sentenceId = 0;
            List <Segment<CellSegmentConfig>> segmentList = new List<Segment<CellSegmentConfig>>();
            int charCount = 0;
            
            foreach (Sheet sheet in sheetList)
            {
                string sheetId = "";
                foreach (var attr in sheet.GetAttributes())
                {
                    if (attr.LocalName == "id")
                    {
                        sheetId = attr.Value;
                        break;
                    }
                }

                WorksheetPart wsPart = (WorksheetPart)(Doc.WorkbookPart.GetPartById(sheetId));
                IEnumerable<Cell> cellList = wsPart.Worksheet.Descendants<Cell>();

                segmentList.Add(new Segment<CellSegmentConfig>() {
                    id  = sentenceId,
                    content = sheet.Name,
                    config = new CellSegmentConfig()
                    {
                        cellName = "",
                        sheet = sheetId
                    }
                });

                ++sentenceId;

                foreach (var cell in cellList)
                {
                        if (cell.CellFormula != null) {
                            continue;
                        }

                        if (cell.DataType != null) {
                            switch (cell.DataType.Value.ToString()) {    
                                case "SharedString":
                                    var stringTable = Doc.WorkbookPart
                                        .GetPartsOfType<SharedStringTablePart>()
                                        .FirstOrDefault();
                                    if (stringTable != null) {
                                        content = stringTable
                                            .SharedStringTable
                                            .ElementAt(int.Parse(cell.InnerText))
                                            .InnerText;
                                    }
                                    break;
                            }
                        } else {
                            content = cell.CellValue != null ? cell.CellValue.Text : null;    
                        }

                        if (String.IsNullOrEmpty(content)) { 
                            continue;
                        }

                        char[] charList = content.ToCharArray();
                        bool isSentenceEnd = false;
                        var prevChar = "";
                        content = "";

                        for (var i = 0; i < charList.Count(); i++) {
                            content = content + charList[i];
                            isSentenceEnd = ServiceStorage.SentenceEndCharList.Contains(charList[i].ToString()) || (
                                charList[i].ToString() == " " &&
                                ServiceStorage.AbstractSentenceEndCharList.Contains(prevChar)
                            );

                            if (isSentenceEnd) {
                                charCount += content.Length + 1;
                                textSegment.content = content;
                                textSegment.id = sentenceId;
                                textSegment.config =  new CellSegmentConfig()
                                {
                                    cellName = cell.GetAttribute("r", "").Value,
                                    sheet = sheetId
                                };;

                                textSegment.hasLineBreak = content.IndexOf("\n") >= 0;
                                segmentList.Add(textSegment);
                                textSegment = new Segment<CellSegmentConfig>();
                                styleList = new List<Style>();
                                content = "";
                                ++sentenceId;
                            }

                            prevChar = charList[i].ToString();
                        }

                        if (!isSentenceEnd) {
                            charCount += content.Length + 1;
                            textSegment.content = content;
                            textSegment.id = sentenceId;
                            textSegment.config = new CellSegmentConfig()
                            {
                                cellName = cell.GetAttribute("r", "").Value,
                                sheet = sheetId   
                            };
                            segmentList.Add(textSegment);
                            content = "";
                            textSegment = new Segment<CellSegmentConfig>();
                            styleList = new List<Style>();
                            ++sentenceId;
                        }
                }
            }

            Doc.Dispose();

            return new Document<CellSegmentConfig>()
            {
                type = Document<CellSegmentConfig>.DocumentType.Spreadsheet,
                segmentList = segmentList,
                charCount = charCount
            };
        }

        public override void SaveDocument(string stream, string filePath, bool isClearStyle)
        {
            List<List<Segment<CellSegmentConfig>>> cellSegmentList 
                = JsonSerializer
                    .Deserialize<List<List<Segment<CellSegmentConfig>>>>(
                        stream, new JsonSerializerOptions() {IncludeFields = true}
                    );

            using (SpreadsheetDocument Doc = SpreadsheetDocument.Open(filePath, true))
            {
                var sharedStringTable = Doc.WorkbookPart
                    .GetPartsOfType<SharedStringTablePart>()
                    .First()
                    .SharedStringTable;

                string content = "";

                Parallel.For(0, cellSegmentList.Count, index =>
                {
                    Segment<CellSegmentConfig> prevSegment = null;

                    int i = 0;
                    foreach (var segment in cellSegmentList.ElementAt(index))
                    {
                        ++i;
                        WorksheetPart wsPart = (WorksheetPart)(Doc.WorkbookPart.GetPartById(segment.config.sheet));
                        if (segment.config.cellName == "")
                        {
                            foreach (Sheet sheet in Doc.WorkbookPart.Workbook.Sheets)
                            {
                                foreach (var attr in sheet.GetAttributes())
                                {
                                    if (attr.LocalName == "id" && attr.Value == segment.config.sheet)
                                    {
                                        sheet.Name = segment.content;
                                    }
                                }
                            }
                        }
                        else
                        {
                            Cell cell = wsPart.Worksheet.Descendants<Cell>()
                                .Where(c => c.CellReference == segment.config.cellName)
                                .FirstOrDefault();

                            if (
                                prevSegment != null  &&
                                segment.config.cellName == prevSegment.config.cellName &&
                                segment.config.sheet == prevSegment.config.sheet
                            )
                            {
                                content += segment.content + ' ';
                            }
                            else
                            {
                                content = segment.content;
                            }

                            if (segment.hasLineBreak)
                            {
                                content += "\n";
                            }
                            
                            ChangeCell(sharedStringTable, cell, content);
                        }

                        prevSegment = segment;
                    }
                });
            }
        }

        private void ChangeCell(
            SharedStringTable sharedStringTable,
            Cell cell, 
            string text
        )
        {
            if (
                cell.DataType == null || cell.DataType != CellValues.SharedString
            )
            {
                cell.RemoveAllChildren();
                cell.AppendChild(new InlineString(new Text { Text = text }));
                cell.DataType = CellValues.InlineString;
                return;
            }
            
            IEnumerable<SharedStringItem> sharedStringItems 
                = sharedStringTable.Elements<SharedStringItem>();
            int i = 0;
            foreach (SharedStringItem sharedStringItem in sharedStringItems)
            {
                if (sharedStringItem.InnerText == text)
                {
                    cell.CellValue = new CellValue(i.ToString());
                    return;
                }
                i++;
            }
            sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            cell.CellValue = new CellValue(i.ToString());
        }
    }
}
