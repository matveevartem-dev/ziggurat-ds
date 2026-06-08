namespace App.Api.Processor.Storage
{
    public interface ServiceStorage
    {
        public const string REPLACED_CRM_CONTENT = "REPLACED_CRM_CONTENT";
        
        public static string[] SentenceEndCharList = {
            "!", "?", "...", ";", "。", "：", "\n", " - ", " — ", ":",
        };

        public static string[] AbstractSentenceEndCharList = {
            "."
        };
    }
}
