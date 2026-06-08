using System.Linq;
using App.Api.Processor.Storage;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace App.Api.Processor.SaveService.Word
{
    public class OxmlFactory: IOpenXmlFactory
    {
        public OpenXmlCompositeElement CreateRun(OpenXmlElement rp)
        {
            return new Run(rp);
        }

        public OpenXmlElement CreateText(string text = "")
        {
            return new Text()
            {
                Text = text,
                Space = SpaceProcessingModeValues.Preserve
            };
        }

        public OpenXmlElement CreateRunProperty()
        {
            return new RunProperties();
        }

        public OpenXmlElement CreateRunProperties()
        {
            return new RunProperties();
        }
    }
}
