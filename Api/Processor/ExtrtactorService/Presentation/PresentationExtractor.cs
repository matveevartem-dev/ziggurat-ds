using System.Collections.Generic;
using System.Linq;
using App.Api.Dto;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Text = DocumentFormat.OpenXml.Drawing.Text;

namespace App.Api.Processor.ExtractorService.Presentation
{
    public class PresentationExtractor : AbstractDocumentExtractor <Hyperlink,
                                                            Text,
                                                            NumberingFormat,
                                                            Run
                                                            >
    {

        public PresentationExtractor(
            PresentationStyleExtractor oxmlStyleExtractor,
            SegmentListCreator segmentListCreator
        )
        {
            OxmlStyleExtractor = oxmlStyleExtractor;
            SegmentListCreator = segmentListCreator;
        }

        public override List<Segment<TextSegmentConfig>> Extract(string filePath)
        {
            List<Segment<TextSegmentConfig>> segmentList = new List<Segment<TextSegmentConfig>>();
            using (PresentationDocument doc = PresentationDocument.Open(filePath, true))
            {
                var slideList = doc.PresentationPart.Presentation.SlideIdList;

                foreach (var item in slideList.Select((value, i) => new {i, value}))
                {
                    // get the "Part" by its "RelationshipId"
                    string slidePartRelationshipId = (item.value as SlideId).RelationshipId;

                    // Get the specified slide part from the relationship ID.
                    SlidePart slidePart =
                        (SlidePart) doc.PresentationPart.GetPartById(slidePartRelationshipId);

                    IEnumerable<Paragraph> pList = slidePart.Slide.Descendants<Paragraph>();

                    SegmentListCreator.Setup(TextSegmentConfig.Location.Document, slidePartRelationshipId);
                    MakeSegmentList(pList);
                }
            }

            return SegmentListCreator.GetSegmentList();
        }

        protected override bool RunHasTabChar(OpenXmlCompositeElement r)
        {
            return r?.PreviousSibling() is Break;
        }
    }
}
