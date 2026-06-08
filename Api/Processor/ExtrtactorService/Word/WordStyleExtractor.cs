using System.Linq;
using App.Api.Dto;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace App.Api.Processor.ExtractorService.Word
{
    public class WordStyleExtractor: AbstractStyleExtractor, IStyleExtractor
    {
        public OpenXmlElement GetStyleBlock(OpenXmlElement run)
        {
            return run.Elements<RunProperties>().FirstOrDefault();
        }

        protected override StyleValue GetBold(OpenXmlElement rp)
        {
            if (rp.Elements<Bold>().FirstOrDefault() != null)
            {
                return MakeBold();
            }

            return null;
        }

        protected override StyleValue GetUnderline(OpenXmlElement rp)
        {
            var underline = rp.Elements<Underline>().FirstOrDefault();
            if (underline != null)
            {
                return MakeUnderline(underline.GetAttributes().First().Value);
            }

            return null;
        }

        protected override StyleValue GetFontSize(OpenXmlElement rp)
        {
            var fontSize = rp.Elements<FontSize>().FirstOrDefault(); 

            if (fontSize != null)
            {
                return MakeFontSize(fontSize.GetAttributes().First().Value);
            }

            return null;
        }

        protected override StyleValue GetColor(OpenXmlElement rp)
        {
            var color = rp.Elements<Color>().FirstOrDefault(); 

            if (color != null)
            {
                return MakeColor(color.GetAttributes().First().Value);
            }

            return null;
        }

        protected override StyleValue GetItalic(OpenXmlElement rp)
        {
            if (rp.Elements<Italic>().FirstOrDefault() != null)
            {
                return MakeItalic();
            }

            return null;
        }

        protected override StyleValue GetHighLight(OpenXmlElement rp)
        {
            var highLight = rp.Elements<Highlight>().FirstOrDefault(); 
            if (highLight != null)
            {
                return MakeHighLight(highLight.GetAttributes().First().Value);
            }

            return null;
        }

        protected override StyleValue GetStrike(OpenXmlElement rp)
        {
            if (rp.Elements<Strike>().FirstOrDefault() != null)
            {
                return MakeStrike();
            }

            return null;
        }

        protected override StyleValue GetColorFill(OpenXmlElement rp)
        {
            RunProperties rrp = (RunProperties) rp;
            if (rrp.Shading != null)
            {
                return MakeColorFill(rrp.Shading.Fill.Value);
            }

            return null;
        }
    }
}
