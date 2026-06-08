using App.Api.Dto;
using App.Api.Processor;
using System;

namespace App.Api.Service;

public class DocumentService
{
    private readonly DocumentFactory _documentFactory;

    public DocumentService(DocumentFactory documentFactory)
    {
        _documentFactory = documentFactory;
    }

    public IDocument ExtractDocument(string filePath)
    {
        var processor = GetRequiredProcessor(filePath);
        return processor.ExtractDocument(filePath);
    }

    // Изменили string на Stream для экономии памяти
    public void SaveDocument(string content, string filePath, bool isClearStyle)
    {
        var processor = GetRequiredProcessor(filePath);
        processor.SaveDocument(content, filePath, isClearStyle);
    }

    private IDocumentProcessor GetRequiredProcessor(string filePath)
    {
        var processor = _documentFactory.GetProcessor(filePath);

        if (processor == null)
        {
            // Используем более универсальное исключение
            throw new NotSupportedException($"File type for '{filePath}' is not supported. Only docx/xlsx/pptx allowed.");
        }

        return processor;
    }
}
