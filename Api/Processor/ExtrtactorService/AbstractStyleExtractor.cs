using System.Collections.Generic;
using System.Linq;
using App.Api.Dto;
using DocumentFormat.OpenXml;
using Style = App.Api.Dto.Style;

namespace App.Api.Processor.ExtractorService
{
    public abstract class AbstractStyleExtractor
    {
        protected abstract StyleValue GetBold(OpenXmlElement rp);
        protected abstract StyleValue GetUnderline(OpenXmlElement rp);
        protected abstract StyleValue GetFontSize(OpenXmlElement rp);
        protected abstract StyleValue GetColor(OpenXmlElement rp);
        protected abstract StyleValue GetItalic(OpenXmlElement rp);
        protected abstract StyleValue GetHighLight(OpenXmlElement rp);
        protected abstract StyleValue GetStrike(OpenXmlElement rp);
        protected abstract StyleValue GetColorFill(OpenXmlElement rp);

        public List<StyleValue> Extract(OpenXmlElement rp)
        {
            List<StyleValue> styleList = new List<StyleValue>();

            styleList.Add(GetBold(rp));
            styleList.Add(GetItalic(rp));
            styleList.Add(GetHighLight(rp));
            styleList.Add(GetColor(rp));
            styleList.Add(GetUnderline(rp));
            styleList.Add(GetStrike(rp));
            styleList.Add(GetFontSize(rp));
            styleList.Add(GetColorFill(rp));
            
            return styleList.Where(s => s != null).ToList();
        }

        protected StyleValue ColorFill()
        {
            return new ()
            {
                type = Style.Type.ColorFill
            };
        }

        protected StyleValue MakeBold()
        {
            return new ()
            {
                type = Style.Type.Bold
            };
        }

        protected StyleValue MakeUnderline(string value)
        { 
            return new ()
            {
                type = Style.Type.Underline,
                value = value
            };
        }

        protected StyleValue MakeFontSize(string value)
        {
            return new ()
            {
                type = Style.Type.Size,
                value = value
            };
        }

        protected StyleValue MakeColor(string value)
        {
            return new ()
            {
                type = Style.Type.Color,
                value = value
            };
        }

        protected StyleValue MakeLumColor(string value)
        {
            return new ()
            {
                type = Style.Type.LumColor,
                value = value
            };
        }

        protected StyleValue MakeItalic()
        {
            return new ()
            {
                type = Style.Type.Italic
            };
        }

        protected StyleValue MakeHighLight(string value)
        {
            return new ()
            {
                type = Style.Type.Highlight,
                value = value
            };
        }

        protected StyleValue MakeStrike()
        {
            return new()
            {
                type = Style.Type.Strike
            };
        }

        protected StyleValue MakeColorFill(string value)
        {
            return new()
            {
                type = Style.Type.ColorFill,
                value = value
            };
        }
    }
}
