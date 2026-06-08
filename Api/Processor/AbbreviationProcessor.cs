using DocumentFormat.OpenXml.Drawing.Charts;
using Microsoft.Extensions.Caching.Memory;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace App.Api.Processor {
    public partial class AbbreviationProcessor
    {
        protected static IMemoryCache Cache;
        protected static Dictionary<int, string> Map = new Dictionary<int, string>();

        // 1. Улучшаем поиск сокращений (буквы и цифры внутри, например м3.)
        // (?<![а-яА-ЯёЁ]) — проверка, что слева нет кириллицы (чтобы не резать обычные слова)
        // ([а-яА-ЯёЁ0-9]{1,4})\. — ищем от 1 до 4 букв или цифр и точку
        private static readonly string AbbrPattern = @"(?<![а-яА-ЯёЁ])([а-яА-ЯёЁ0-9]{1,4})\.";



        public AbbreviationProcessor(IMemoryCache cache)
        {
            Cache = cache;
        }

        /// <summary>
        /// Кодирует сокращения в теги {{#HASH#}} и сохраняет оригиналы в кэш.
        /// </summary>
        public string Encode(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;

            return Regex.Replace(input, AbbrPattern, m =>
            {
                string original = m.Value; // Здесь будет "м2."
                string cacheKey = GetUniqueHash(original);
                GetOrSetData(cacheKey, original);

                return $"{{#{cacheKey}#}}";
            });
        }

        public string Encode1(string input)
        {
            return Regex.Replace(input, AbbrPattern, m =>
            {
                string hashKey = GetUniqueHash(m.Value);
                GetOrSetData(hashKey, m.Value);
                return $"{{#{hashKey}#}}"; // Возвращаем ТОЛЬКО тег, он сам заменит m.Value
            });
        }

        // 2. Исправляем Decode (теперь ищет именно то, что создал Encode)
        public string Decode(string input)
        {
            return AbbrRegex().Replace(input, m =>
            {
                // Извлекаем HASH из первой группы захвата (...)
                string hashFromTag = m.Groups[1].Value;

                string cachedValue = GetOrSetData(hashFromTag, null);

                // Если нашли в кэше — возвращаем оригинал, иначе оставляем тег
                return !string.IsNullOrEmpty(cachedValue) ? cachedValue : m.Value;
            });
        }

        public static string GetData(string hash)
        {
            // Пытаемся получить данные из кэша
            if (!AbbreviationProcessor.Cache.TryGetValue(hash, out string value))
            {
                // Если данных нет, имитируем "тяжелый" запрос (например, из БД)
                value = "Результат из базы данных: " + DateTime.Now;

                // Настройки хранения
                var cacheOptions = new MemoryCacheEntryOptions()
                    .SetAbsoluteExpiration(TimeSpan.FromMinutes(10)) // Удалить через 10 минут точно
                    .SetSlidingExpiration(TimeSpan.FromMinutes(2));  // Удалить, если к данным не обращались 2 минуты

                // Сохраняем в кэш
                AbbreviationProcessor.Cache.Set(hash, value, cacheOptions);
            }

            return value;
        }

        public static string GetOrSetData(string hash, string? input = null)
        {
            // 1. Проверяем кэш
            if (Cache.TryGetValue(hash, out object? cachedData))
            {
                // Если нашли, возвращаем данные из кэша
                return cachedData?.ToString() ?? string.Empty;
            }

            // 2. Если в кэше нет, но пришел новый input — сохраняем его
            if (input != null)
            {
                var cacheOptions = new MemoryCacheEntryOptions()
                    .SetAbsoluteExpiration(TimeSpan.FromMinutes(10))
                    .SetSlidingExpiration(TimeSpan.FromMinutes(2));

                Cache.Set(hash, input, cacheOptions);
                return input;
            }

            return string.Empty;
        }


        public static string SetData(string input)
        {
            var cacheOptions = new MemoryCacheEntryOptions()
                .SetAbsoluteExpiration(TimeSpan.FromMinutes(10)) // Удалить через 10 минут точно
                .SetSlidingExpiration(TimeSpan.FromMinutes(2));  // Удалить, если к данным не обращались 2 минуты

            return Cache.Set(GetUniqueHash(input), input, cacheOptions);
        }

        public static string GetUniqueHash(string input)
        {
            using (SHA256 sha256Hash = SHA256.Create())
            {
                byte[] bytes = sha256Hash.ComputeHash(Encoding.UTF8.GetBytes(input));
                return Convert.ToHexString(bytes); // Returns a 64-character unique string
            }
        }

        // 3. Синхронизируем количество скобок с методом Encode {{# ... #}}
        [GeneratedRegex(@"{#([A-F0-9]{64})#}", RegexOptions.IgnoreCase)]
        private static partial Regex AbbrRegex();
    }
}
