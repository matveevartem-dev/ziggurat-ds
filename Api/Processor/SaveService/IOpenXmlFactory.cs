
using System.Collections.Generic;
using App.Api.Dto;
using DocumentFormat.OpenXml;

namespace App.Api.Processor.SaveService
{
    public interface IOpenXmlFactory
    {
        public OpenXmlCompositeElement CreateRun(OpenXmlElement rp);
        public OpenXmlElement CreateText(string text = "");
        public OpenXmlElement CreateRunProperty();
    }
}
