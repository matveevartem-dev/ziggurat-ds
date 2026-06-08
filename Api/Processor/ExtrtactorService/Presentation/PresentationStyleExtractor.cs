using System.Linq;
using App.Api.Dto;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;

namespace App.Api.Processor.ExtractorService.Presentation
{
    public class PresentationStyleExtractor: AbstractStyleExtractor, IStyleExtractor
    {
        public OpenXmlElement GetStyleBlock(OpenXmlElement run)
        {
            return run.Elements<RunProperties>().FirstOrDefault();
        }

        protected override StyleValue GetBold(OpenXmlElement rp)
        {
            var bold = rp.GetAttributes().FirstOrDefault(a => a.LocalName == "b");
            return bold.Value != null ? MakeBold() : null;
        }

        protected override StyleValue GetUnderline(OpenXmlElement rp)
        {
            var underline = rp.GetAttributes().FirstOrDefault(a => a.LocalName == "u");
            return underline.Value != null ? MakeUnderline(underline.Value) : null;
        }

        protected override StyleValue GetFontSize(OpenXmlElement rp)
        {
            var fontSize = rp.GetAttributes().FirstOrDefault(a => a.LocalName == "sz");
            return fontSize.Value != null ? MakeFontSize(fontSize.Value) : null;
        }

        protected override StyleValue GetColor(OpenXmlElement rp)
        {
            var color = rp.Elements<SolidFill>().FirstOrDefault();
            if (color?.RgbColorModelHex != null)
            {
                return MakeColor(color.RgbColorModelHex.Val.Value);
            }

            LuminanceModulation lumMode = color?.SchemeColor?.Elements<LuminanceModulation>().FirstOrDefault();
            
            if (lumMode != null)
            {
                return MakeLumColor(lumMode.Val);
            }

            return null;
        }

        protected override StyleValue GetItalic(OpenXmlElement rp)
        {
            var italic = rp.GetAttributes().FirstOrDefault(a => a.LocalName == "i");

            return italic.Value != null ? MakeItalic() : null;
        }

        protected override StyleValue GetHighLight(OpenXmlElement rp)
        {
            Highlight highLight = rp.Elements<Highlight>().FirstOrDefault(); 
            if (highLight != null)
            {
                return MakeHighLight(highLight.RgbColorModelHex.Val);
            }

            return null;
        }

        protected override StyleValue GetStrike(OpenXmlElement rp)
        {
            var strike = rp.GetAttributes().FirstOrDefault(a => a.LocalName == "strike");

            return strike.Value != null ? MakeStrike() : null;
        }

        protected override StyleValue GetColorFill(OpenXmlElement rp)
        {
            return null;
        }
    }
}
