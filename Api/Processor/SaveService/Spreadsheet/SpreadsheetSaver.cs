using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using App.Api.Dto;
using App.Api.Processor.ExtractorService.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace App.Api.Processor.SaveService.Spreadsheet
{
    public class SpreadsheetSaver:  ISaveService<CellSegmentConfig>
    {
        private OxmlFactory _oxmlFactory;
        private SpreadsheetStyleExtractor _oxmlExtractor;
        private StyleSaver _styleSaver;
            
        public SpreadsheetSaver(
            SpreadsheetStyleExtractor oxmlExtractor,
            StyleSaver styleSaver,
            OxmlFactory oxmlFactory
            )
        {
            _oxmlFactory = oxmlFactory;
            _oxmlExtractor = oxmlExtractor;
            _styleSaver = styleSaver;
        }
        
        private IEnumerable<SharedStringItem> getCellList(
            TextSegmentConfig.Location location,
            WordprocessingDocument doc,
            int pageId
        )
        {

            return null;

        }
        
        protected void RemoveText(OpenXmlCompositeElement run)
        {
            
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

        public void Save(List<List<Segment<CellSegmentConfig>>> cellSegmentList, string filePath, bool isClearStyle)
        {
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
    }
}
