using App.Api.Dto;
using System.Collections.Generic;

namespace App.Api.Processor
{
    public interface IDocumentProcessor
    {
        public IDocument ExtractDocument(string filePath);
        public void SaveDocument(string doc, string filePath, bool isClearStyle);
    }
}