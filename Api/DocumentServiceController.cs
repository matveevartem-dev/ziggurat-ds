using App.Api.Dto;
using App.Api.Service;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.IO;
using App.Api.Extensions;
using System.Linq;

namespace App.Api
{
    [ApiController]
    [Route("[controller]")]
    public class DocumentServiceController : ControllerBase
    {
        private readonly DocumentService _documentService;
        private readonly IMemoryCache _cache; // Поле для кэша

        public DocumentServiceController(DocumentService documentService, IMemoryCache cache)
        {
            _documentService = documentService;
            _cache = cache;
        }

        /// <summary>
        /// Получает экземпляр IDocument из файла
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException"></exception>
        /*[HttpGet]
        public IDocument Get(string filePath)
        {
            if (false == System.IO.File.Exists(filePath)) {
                throw new FileNotFoundException(filePath);
            }
            
            return _documentService.ExtractDocument(filePath);
        }*/

        [HttpGet]
        public IDocument Get(string filePath)
        {
            // 1. Пытаемся достать из кэша
            /*if (_cache.TryGetValue(filePath, out IDocument cachedDoc))
            {
                return cachedDoc;
            }*/

            // 2. Если в кэше нет, проверяем файл
            if (!System.IO.File.Exists(filePath))
            {
                throw new FileNotFoundException(filePath);
            }

            // 3. Извлекаем документ через сервис
            var document = _documentService.ExtractDocument(filePath);

            // 4. Сохраняем в кэш на будущее (например, на 10 минут)
            //_cache.Set(filePath, document, TimeSpan.FromMinutes(10));

            return document;
        }

        /// <summary>
        /// Сохранение в файл 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="clearStyle"></param>
        /// <exception cref="FileNotFoundException"></exception>
        [HttpPost]
        [Produces("application/json")]
        public void Post(string filePath, bool clearStyle = false)
        {
            if (false == System.IO.File.Exists(filePath)) {
                throw new FileNotFoundException(filePath);
            }

            var rawMessage = new StreamReader(Request.Body)
                .ReadToEndAsync()
                .Result;
            _documentService.SaveDocument(rawMessage, filePath, clearStyle);
            Console.WriteLine('!');
        }

        /// <summary>
        /// Посмотреть список всех ключей в кэше (для отладки)
        /// </summary>
        [HttpGet("debug/cache")]
        public IActionResult GetCacheStatus()
        {
            // Используем наш метод расширения
            var keys = MemoryCacheExtensions.GetKeys(_cache);

            var report = keys.Select(k => new {
                Key = k.ToString(),
                // Можно попробовать вытащить само значение, если нужно
                Exists = _cache.TryGetValue(k, out var value),
                Type = value?.GetType().Name
            });

            return Ok(report);
        }
    }
}
