using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using App.Api.Dto;
using App.Api.Processor.Storage;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace App.Api.Processor.ExtractorService.Spreadsheet
{
    public class SpreadsheetExtractor: IExtractorService<CellSegmentConfig>
    {
        private SpreadsheetStyleExtractor _oxmlStyleExtractor;

        public SpreadsheetExtractor(SpreadsheetStyleExtractor oxmlStyleExtractor)
        {
            _oxmlStyleExtractor = oxmlStyleExtractor;
        }

        public List<Segment<CellSegmentConfig>> Extract(string filePath)
        {
            SpreadsheetDocument Doc = SpreadsheetDocument.Open(filePath, false);
            Segment<CellSegmentConfig> textSegment = new Segment<CellSegmentConfig>();
            List<Style> styleList = new List<Style>();
            List<Segment<CellSegmentConfig>> segmentList = new List<Segment<CellSegmentConfig>>();
            var sheetList = Doc.WorkbookPart.Workbook.Sheets;
            string content = null;
            int sentenceId = 0;
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

                segmentList.Add(
                        new Segment<CellSegmentConfig>() 
                        {
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

                    StyleIterator styleIterator = new StyleIterator();
                    if (cell.DataType != null) {
                        switch (cell.DataType.Value.ToString()) {
                            case "SharedString":
                                var stringTable = Doc.WorkbookPart
                                    .GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();
                                if (stringTable != null)
                                {
                                    content = stringTable
                                        .SharedStringTable
                                        .ElementAt(int.Parse(cell.InnerText))
                                        .InnerText;
                                    
                                    styleIterator = GetStyleIterator(
                                        stringTable.SharedStringTable
                                            .ElementAt(int.Parse(cell.InnerText)),
                                            sentenceId
                                    );
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
                            textSegment.styleList = styleIterator.GetStyleList();
                            textSegment.config =  new CellSegmentConfig()
                            {
                                cellName = cell.GetAttribute("r", "").Value,
                                sheet = sheetId
                            };

                            textSegment.hasLineBreak = content.IndexOf("\n") >= 0;
                            segmentList.Add(textSegment);

                            textSegment = new Segment<CellSegmentConfig>();
                            content = "";
                            ++sentenceId;
                        }

                        prevChar = charList[i].ToString();
                    }

                    if (!isSentenceEnd) {
                        charCount += content.Length + 1;
                        textSegment.content = content;
                        textSegment.id = sentenceId;
                        textSegment.styleList = styleIterator.GetStyleList();
                        textSegment.config = new CellSegmentConfig()
                        {
                            cellName = cell.GetAttribute("r", "").Value,
                            sheet = sheetId
                        };
                        segmentList.Add(textSegment);
                        content = "";
                        textSegment = new Segment<CellSegmentConfig>();
                        ++sentenceId;
                    }
                }
            }

            Doc.Dispose();

            return segmentList;
        }

        private StyleIterator GetStyleIterator(
            OpenXmlElement cell,
            int sentenceId
        )
        {
            StyleIterator styleIterator = new StyleIterator();
            string oldStyleList = null;

            List<Run> rList = cell.Elements<Run>().ToList();
            string content = "";
            foreach (var r in rList.Select((value, i) => new {i, value}))
            {
                OpenXmlElement rp = _oxmlStyleExtractor.GetStyleBlock(cell);
                if (rp != null && r.value.InnerText.Trim() != "")
                {
                    List<StyleValue> styleValueList = _oxmlStyleExtractor.Extract(rp);
                    bool hasStyle = styleValueList.Count > 0;

                    string serializedStyleList = JsonSerializer.Serialize(styleValueList);

                    bool isOldStyleTheSame = serializedStyleList == oldStyleList;
                    if (isOldStyleTheSame && styleIterator.IsNotEmpty())
                    {
                        styleIterator.Last().endPos = content.Length + r.value.InnerText.Length;
                    }

                    if (hasStyle && !isOldStyleTheSame)
                    {
                        Style style = new Style() {
                            segmentId = sentenceId,
                            startPos = content.Length,
                            endPos = content.Length + r.value.InnerText.Length,
                            value = rp.InnerXml,
                            styleValueList = styleValueList,
                        };
                        styleIterator.Add(style);
                    }

                    oldStyleList = JsonSerializer.Serialize(styleValueList);
                    content += r.value.InnerText;
                }
            }

            return styleIterator;
        }

        public int CalculateContentLength(List<Segment<CellSegmentConfig>> segmentList)
        {
            int contentLength = 0;
            foreach (Segment<CellSegmentConfig> segment in segmentList)
            {
                contentLength = contentLength + segment.content.Length;
            }

            return contentLength;
        }

        public OpenXmlElement GetStyleBlock(OpenXmlElement cell)
        {
            SharedStringItem ssi = (SharedStringItem) cell;
            return ssi.Elements<RunProperties>().FirstOrDefault();
        }

        public OpenXmlElement GetParentHyperlink(OpenXmlCompositeElement run)
        {
            return null;
        }
        
        public List<OpenXmlCompositeElement> GetRunList(OpenXmlCompositeElement p)
        {
            return null;
        }
    }
}
