
using App.Api.Dto;
using App.Api.Processor.ExtractorService;
using App.Api.Processor.SaveService;

namespace App.Api.Processor
{
    public abstract class AbstractDocumentProcessor<T>: IDocumentProcessor
    {
        protected IExtractorService<T> _extractorService;
        protected ISaveService<TextSegmentConfig> _saveService;

        public abstract IDocument ExtractDocument(string filePath);
        public abstract void SaveDocument(string doc, string filePath, bool isClearStyle);
    }
}
