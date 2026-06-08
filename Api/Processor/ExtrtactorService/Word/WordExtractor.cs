using System;
using System.Collections.Generic;
using System.Linq;
using App.Api.Dto;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace App.Api.Processor.ExtractorService.Word
{
    public class WordExtractor : AbstractDocumentExtractor <Hyperlink,
                                                            Text,
                                                            NumberingProperties,
                                                            Run
                                                            >
    {
        public WordExtractor(
            WordStyleExtractor oxmlStyleExtractor,
            SegmentListCreator segmentListCreator
        )
        {
            OxmlStyleExtractor = oxmlStyleExtractor;
            SegmentListCreator = segmentListCreator;
        }

        protected override bool RunHasTabChar(OpenXmlCompositeElement r)
        {
            if (r == null)
            {
                return false;
            }

            return r.Descendants<TabChar>().FirstOrDefault() != null || 
                   r.Descendants<Break>().FirstOrDefault() != null ||
                   r.Descendants<Text>().FirstOrDefault() == null;
        }

        public override List<Segment<TextSegmentConfig>> Extract(string filePath)
        {
            List<Segment<TextSegmentConfig>> segmentList = new List<Segment<TextSegmentConfig>>();
            using (WordprocessingDocument Doc = WordprocessingDocument.Open(filePath, true))
            {
                IEnumerable<Paragraph> pdList = Doc.MainDocumentPart.Document.Body.Descendants<Paragraph>().ToList();
                IEnumerable<Paragraph> pfList = new List<Paragraph>();
                IEnumerable<Paragraph> pHeaderlist = new List<Paragraph>();
                IEnumerable<Paragraph> pFooterlist = new List<Paragraph>();
                FootnotesPart footnotes = Doc.MainDocumentPart.FootnotesPart;
                IEnumerable<HeaderPart> headerList = Doc.MainDocumentPart.HeaderParts;
                IEnumerable<FooterPart> footerList = Doc.MainDocumentPart.FooterParts;
                
                if (footnotes != null)
                {
                    pfList = footnotes.Footnotes.Descendants<Paragraph>().ToList();
                }

                SegmentListCreator.Setup(TextSegmentConfig.Location.Document);
                MakeSegmentList(pdList);

                SegmentListCreator.Setup(TextSegmentConfig.Location.Footnotes);
                MakeSegmentList(pfList);

                for (int index = 0; index < headerList.Count(); index++)
                {
                    pHeaderlist = headerList.ElementAt(index).RootElement.Descendants<Paragraph>();
                    SegmentListCreator.Setup(TextSegmentConfig.Location.Header, (index + 1).ToString());
                    MakeSegmentList(pHeaderlist);
                }

                for (int index = 0; index < footerList.Count(); index++)
                {
                    pFooterlist = footerList.ElementAt(index).RootElement.Descendants<Paragraph>();
                    SegmentListCreator.Setup(TextSegmentConfig.Location.Footer, (index + 1).ToString());
                    MakeSegmentList(pFooterlist);
                }

                Console.WriteLine(CalculateContentLength(segmentList));
            }

            return SegmentListCreator.GetSegmentList();
        }
    }
}
