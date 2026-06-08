using System.IO;

namespace App.Api.Processor
{
    public class DocumentFactory
    {
        private readonly WordProcessor _wordProcessor;
        private readonly SpreadsheetProcessor _spreadsheetProcessor;
        private readonly PresentationProcessor _presentationProcessor;

        public DocumentFactory(
            WordProcessor wordProcessor,
            SpreadsheetProcessor spreadsheetProcessor,
            PresentationProcessor presentationProcessor
        )
        {
            _wordProcessor = wordProcessor;
            _spreadsheetProcessor = spreadsheetProcessor;
            _presentationProcessor = presentationProcessor;
        }

        public IDocumentProcessor GetProcessor(string filePath)
        {
            switch (Path.GetExtension(filePath)) {
                case ".docx":
                    return _wordProcessor;
                case ".xlsx": 
                    return _spreadsheetProcessor;
                case ".pptx":
                    return _presentationProcessor;
                default:
                    return null;
            }
        }
    }
}
