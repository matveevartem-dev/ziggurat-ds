using App.Api.Dto;
using DocumentFormat.OpenXml;

namespace App.Api.Processor.SaveService
{
    public interface IStyleSaver
    {
        public void Save(StyleValue styleValue, OpenXmlElement rp);
        public void RemoveStyles(OpenXmlElement rp);
    }
}
