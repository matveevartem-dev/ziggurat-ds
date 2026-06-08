using App.Api.Dto;
using App.Api.Processor;
using App.Api.Processor.ExtractorService.Presentation;
using App.Api.Processor.ExtractorService.Word;
using App.Api.Processor.SaveService.Presentation;
using App.Api.Processor.SaveService.Word;
using App.Api.Service;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.OpenApi;
//using Microsoft.OpenApi.Models; // Исправлено для OpenApiInfo
using PresentationElementCreator = App.Api.Processor.SaveService.Presentation.ElementCreator;
// Алиасы для разрешения конфликтов имен
using PresentationStyleSaver = App.Api.Processor.SaveService.Presentation.StyleSaver;
using WordStyleSaver = App.Api.Processor.SaveService.Word.StyleSaver;

namespace App;

public static class DependencyInjection
{
    public static IServiceCollection AddApplicationServices(this IServiceCollection services)
    {
        //services.AddControllers();
        //services.AddMemoryCache();

        // Общие сервисы
        services.AddScoped<DocumentService>();
        services.AddScoped<DocumentFactory>(); // Вероятно, может быть Singleton
        services.AddScoped<OxmlFactory>();     // Вероятно, может быть Singleton

        // Word
        services.AddWordServices();

        // Presentation
        services.AddPresentationServices();

        // Spreadsheet
        services.AddScoped<SpreadsheetProcessor>();
        services.AddScoped<SegmentListCreator>();
        services.AddScoped<App.Api.Processor.AbbreviationProcessor>();

        return services;
    }

    private static void AddControllers(this IServiceCollection services)
    {
    }

    private static void AddWordServices(this IServiceCollection services)
    {
        services.AddScoped<WordProcessor>();
        services.AddScoped<WordExtractor>();
        services.AddScoped<WordSaver>();
        services.AddScoped<WordStyleExtractor>();
        services.AddScoped<WordStyleSaver>();
    }

    private static void AddPresentationServices(this IServiceCollection services)
    {
        services.AddScoped<PresentationProcessor>();
        services.AddScoped<PresentationExtractor>();
        services.AddScoped<PresentationSaver>();
        services.AddScoped<PresentationStyleExtractor>();
        services.AddScoped<PresentationStyleSaver>();
        services.AddScoped<PresentationElementCreator>();
    }

    public static IServiceCollection AddSwaggerDocumentation(this IServiceCollection services)
    {
        services.AddEndpointsApiExplorer();
        services.AddSwaggerGen(c =>
        {
            c.SwaggerDoc("v1", new OpenApiInfo { Title = "Document service", Version = "v1" });
        });
        return services;
    }
}
