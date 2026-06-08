
using System.Collections.Generic;
using App.Api.Dto;
using DocumentFormat.OpenXml;

namespace App.Api.Processor.ExtractorService
{
    public interface IStyleExtractor
    {
        public OpenXmlElement GetStyleBlock(OpenXmlElement run);

        public List<StyleValue> Extract(OpenXmlElement rp);
    }
}
