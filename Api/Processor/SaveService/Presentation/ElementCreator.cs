using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace App.Api.Processor.SaveService.Presentation
{
    public class ElementCreator: IOpenXmlFactory
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
            };
        }

        public OpenXmlElement CreateRunProperty()
        {
            return new RunProperties();
        }

        public OpenXmlElement CreateRunPropertiesStyle(OpenXmlElement rp, bool isClear = false)
        {
            return (RunProperties) rp.CloneNode(true);
        }
    }
}
