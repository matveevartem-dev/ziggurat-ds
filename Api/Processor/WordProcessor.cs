using App.Api.Dto;
using System.Collections.Generic;
using System.Text.Json;
using App.Api.Processor.ExtractorService.Word;
using App.Api.Processor.SaveService;
using App.Api.Processor.SaveService.Word;

namespace App.Api.Processor
{
    public class WordProcessor: AbstractDocumentProcessor<TextSegmentConfig>
    {
        public WordProcessor(
            WordExtractor extractorService,
            WordSaver saveService
        )
        {
            _extractorService = extractorService;
            _saveService = (ISaveService<TextSegmentConfig>) saveService;
        }

        /// <summary>
        /// Парсер DOCX файла
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public override IDocument ExtractDocument(string filePath)
        { 
            WordFixer.Fix(filePath);

            List <Segment<TextSegmentConfig>> segmentList = _extractorService.Extract(filePath);
            
            return new Document<TextSegmentConfig>()
            {
                type = Document<TextSegmentConfig>.DocumentType.Text,
                segmentList = segmentList,
                charCount = _extractorService.CalculateContentLength(segmentList)
            };
        }

        /// <summary>
        /// Cохраняет в DOCX файл JSON массив из сегментов,
        /// объеденённых в параграфы.
        /// </summary>
        /// <param name="stream">Массив из наборов сегментов типа App.Api.Dto.Segment[][]</param>
        /// <param name="filePath">Абсолюбтный путь к файлу</param>
        public override void SaveDocument(string stream, string filePath, bool isClearStyle)
        {
            List<List<Segment<TextSegmentConfig>>> paragraphSegmentList 
                = JsonSerializer
                    .Deserialize<List<List<Segment<TextSegmentConfig>>>>(
                        stream, new JsonSerializerOptions() {IncludeFields = true}
                        );

            _saveService.Save(paragraphSegmentList, filePath, isClearStyle);
        }
    }
}
