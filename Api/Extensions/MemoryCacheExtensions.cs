namespace App.Api.Extensions
{
    using Microsoft.Extensions.Caching.Memory;
    using System.Collections.Generic;
    using System.Linq;
    using System.Reflection;

    public static class MemoryCacheExtensions
    {
        /*public static IEnumerable<object> GetKeys(this IMemoryCache cache) // Добавили this
        {
            var coherentStateField = typeof(MemoryCache).GetField("_coherentState", BindingFlags.NonPublic | BindingFlags.Instance);
            var coherentState = coherentStateField?.GetValue(cache);
            var entriesProperty = coherentState?.GetType().GetProperty("Entries", BindingFlags.Public | BindingFlags.Instance);
            var entries = entriesProperty?.GetValue(coherentState) as System.Collections.IDictionary;

            return entries?.Keys.Cast<object>() ?? Enumerable.Empty<object>();
        }*/

        public static IEnumerable<object> GetKeys(IMemoryCache cache)
        {
            var coherentStateField = typeof(MemoryCache).GetField("_coherentState", BindingFlags.NonPublic | BindingFlags.Instance);
            var coherentState = coherentStateField?.GetValue(cache);
            var entriesProperty = coherentState?.GetType().GetProperty("Entries", BindingFlags.Public | BindingFlags.Instance);
            var entries = entriesProperty?.GetValue(coherentState) as System.Collections.IDictionary;

            return entries?.Keys.Cast<object>() ?? Enumerable.Empty<object>();
        }
    }
}
