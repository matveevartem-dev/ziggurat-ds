using App.Api.Dto;
using System.Collections.Generic;
using System.Text.Json;
using App.Api.Processor.ExtractorService.Presentation;
using App.Api.Processor.SaveService;
using App.Api.Processor.SaveService.Presentation;

namespace App.Api.Processor
{
    public class PresentationProcessor : AbstractDocumentProcessor<TextSegmentConfig>
    {
        public PresentationProcessor(
            PresentationExtractor extractorService,
            PresentationSaver saveService
            )
        {
            _extractorService = extractorService;
            _saveService = (ISaveService<TextSegmentConfig>) saveService;
        }

        public override IDocument ExtractDocument(string filePath)
        {
            List<Segment<TextSegmentConfig>> segmentList = _extractorService.Extract(filePath);
            
            return new Document<TextSegmentConfig>()
            {
                type = Document<TextSegmentConfig>.DocumentType.Text,
                segmentList = segmentList,
                charCount = _extractorService.CalculateContentLength(segmentList)
            };
        }

        public override void SaveDocument(string doc, string filePath, bool isClearStyle)
        {
            List<List<Segment<TextSegmentConfig>>> paragraphSegmentList
                = JsonSerializer
                    .Deserialize<List<List<Segment<TextSegmentConfig>>>>(
                        doc, new JsonSerializerOptions() {IncludeFields = true}
                    );

            _saveService.Save(paragraphSegmentList, filePath, isClearStyle);
        }
    }
}
