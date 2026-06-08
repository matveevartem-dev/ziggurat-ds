using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

namespace App.Api.Extensions;

public static class ServiceExtensions
{
    // Метод для регистрации сервисов (аналог ConfigureServices)
    public static IServiceCollection AddAppServices(this IServiceCollection services, IConfiguration config)
    {
        services.AddControllers();
        // Пример добавления своих сервисов:
        // services.AddScoped<IMyService, MyService>();

        return services;
    }

    // Метод для настройки Middleware (аналог Configure)
    public static WebApplication UseAppMiddleware(this WebApplication app)
    {
        if (app.Environment.IsDevelopment())
        {
            app.UseDeveloperExceptionPage();
        }

        app.UseHttpsRedirection();
        app.UseAuthorization();
        app.MapControllers();

        return app;
    }
}
