
using System.Collections.Generic;
using App.Api.Dto;
using DocumentFormat.OpenXml;

namespace App.Api.Processor.ExtractorService
{
    public interface IExtractorService<T>
    {
        //"!", "?", "...", ";", "。", "：", "\n", " - ", " — ", ":",
        public static string[] SentenceEndCharList = {
            "@"
        };

        public static string[] AbstractSentenceEndCharList = {
            "~"
        };

        public List<Segment<T>> Extract(string filePath);

        public int CalculateContentLength(List<Segment<T>> segmentList);
        
        public OpenXmlElement GetStyleBlock(OpenXmlElement run);

        public List<OpenXmlCompositeElement> GetRunList(OpenXmlCompositeElement p);

        public OpenXmlElement GetParentHyperlink(OpenXmlCompositeElement run);
    }
}
