using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using App.Api.Dto;
using App.Api.Processor.ExtractorService.Presentation;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace App.Api.Processor.SaveService.Presentation
{
    public class PresentationSaver: AbstractDocumentSaver, ISaveService<TextSegmentConfig>
    {
        
        public PresentationSaver(
            PresentationExtractor oxmlExtractor,
            StyleSaver styleSaver,
            ElementCreator oxmlFactory
            )
        {
            _oxmlFactory = oxmlFactory;
            _styleSaver = styleSaver;
            _oxmlExtractor = oxmlExtractor;
        }

        public void Save(List<List<Segment<TextSegmentConfig>>> paragraphSegmentList, string filePath, bool isClearStyle)
        {
            using (PresentationDocument doc = PresentationDocument.Open(filePath, true))
            {
                foreach (var segmentBlock in paragraphSegmentList)
                {
                    Segment<TextSegmentConfig> firstSegment = segmentBlock.First();
                    string slidePartRelationshipId = firstSegment.config.pageId;
                    SlidePart slidePart =
                        (SlidePart) doc.PresentationPart.GetPartById(slidePartRelationshipId);
                    
                    List<Paragraph> pList = slidePart.Slide.Descendants<Paragraph>().ToList();
                    Parallel.For(0, paragraphSegmentList.Count,
                        index =>
                        {
                            SaveParagraph(paragraphSegmentList, index, pList, isClearStyle);
                        });
                }
            }
        }

        protected override void RemoveText(OpenXmlCompositeElement run)
        {
            Text text = run.Elements<Text>().FirstOrDefault();
            
            if (text != null)
            {
                text.Text = "";
            }
        }
    }
}
