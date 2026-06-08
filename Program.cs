using App;
using App.Api.Extensions;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;


var builder = WebApplication.CreateBuilder(args);

// 1. Регистрируем ваши сервисы приложения
builder.Services.AddControllers();
builder.Services.AddMemoryCache();
builder.Services.AddApplicationServices();

// 2. Регистрируем Swagger (ваш второй метод)
builder.Services.AddSwaggerDocumentation();

var app = builder.Build();

// Настройка Middleware (конвейера)
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

//app.UseHttpsRedirection();
//app.UseAuthorization();

app.MapControllers();
app.Run();
