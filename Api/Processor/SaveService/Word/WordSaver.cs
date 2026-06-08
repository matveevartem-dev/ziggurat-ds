using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using App.Api.Dto;
using App.Api.Processor.ExtractorService.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace App.Api.Processor.SaveService.Word
{
    public class WordSaver: AbstractDocumentSaver, ISaveService<TextSegmentConfig>
    {
        public WordSaver(
            WordExtractor oxmlExtractor,
            StyleSaver styleSaver,
            OxmlFactory oxmlFactory
        )
        {
            _oxmlFactory = oxmlFactory;
            _oxmlExtractor = oxmlExtractor;
            _styleSaver = styleSaver;
        }

        public void Save(List<List<Segment<TextSegmentConfig>>> paragraphSegmentList, string filePath, bool isClearStyle)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
            {
                Segment<TextSegmentConfig> firstSegment = paragraphSegmentList.First().First();
                var pageId = 0;
                if (firstSegment.config.pageId != null && firstSegment.config.pageId != "")
                {
                    pageId = Int16.Parse(firstSegment.config.pageId);
                }

                IEnumerable<Paragraph> pList = getParagraphList(
                    firstSegment.config.location, 
                    doc,
                    pageId
                );

                Parallel.For(0, paragraphSegmentList.Count, index => {
                    SaveParagraph(paragraphSegmentList,  index, pList, isClearStyle);
                });
            }
        }

        private IEnumerable<Paragraph> getParagraphList(
            TextSegmentConfig.Location location,
            WordprocessingDocument doc,
            int pageId
        )
        {
            IEnumerable<Paragraph> pList = new List<Paragraph>();

            if (location == TextSegmentConfig.Location.Document)
            {
                pList = doc.MainDocumentPart.Document.Body.Descendants<Paragraph>();    
            }
            else if (location == TextSegmentConfig.Location.Footnotes)
            {
                pList = doc.MainDocumentPart.FootnotesPart.Footnotes.Descendants<Paragraph>(); 
            }
            else if (location == TextSegmentConfig.Location.Header)
            {
                pList = doc.MainDocumentPart.HeaderParts.ElementAt(pageId - 1).RootElement.Descendants<Paragraph>();
            }
            else if (location == TextSegmentConfig.Location.Footer)
            {
                pList = doc.MainDocumentPart.FooterParts.ElementAt(pageId - 1).RootElement.Descendants<Paragraph>();
            }

            return pList;
        }

        protected override void RemoveText(OpenXmlCompositeElement run)
        {
            // Text text = run.Elements<Text>().FirstOrDefault();
            foreach (var text in  run.Elements<Text>())
            {
                text.Text = "";
            }

            return;
            /*var vanish = new Vanish() {Val = true};
            RunProperties rp = run.Elements<RunProperties>().FirstOrDefault();
            if (rp == null)
            {
                rp = new RunProperties();
                run.Append(rp);
            }

            rp?.Append(vanish);*/
        }
    }
}
