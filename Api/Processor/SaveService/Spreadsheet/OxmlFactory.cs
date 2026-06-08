using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace App.Api.Processor.SaveService.Spreadsheet
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
    }
}
