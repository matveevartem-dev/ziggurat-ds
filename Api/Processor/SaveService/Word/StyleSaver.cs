using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace App.Api.Processor.SaveService.Word
{
    public class StyleSaver: AbstractStyleSaver
    {
        protected override void SaveBold(OpenXmlElement rp)
        {
            Bold bold = new Bold();
            rp?.Append(bold);
        }

        protected override void SaveItalic(OpenXmlElement rp)
        {
            Italic italic = new Italic();
            rp.Append(italic);
        }

        protected override void SaveHighlight(OpenXmlElement rp, string color)
        {
            color = char.ToUpper(color[0]) + color.Substring(1);
            Highlight highlight = new Highlight()
            {
                Val = (HighlightColorValues) Enum.Parse(typeof(HighlightColorValues), color)
            };
            rp.Append(highlight);
        }

        protected override void SaveUnderline(OpenXmlElement rp, string style)
        {
            UnderlineValues value = UnderlineValues.None;
            
            if (style != "none")
            {
                value = UnderlineValues.Single;
            }

            Underline underline = new Underline() {Val = value};
            rp.Append(underline);
        }

        protected override void SaveColor(OpenXmlElement rp, string colorCode)
        {
            Color color = new Color() {Val = colorCode};
            rp.Append(color);
        }

        protected override void SaveLumColor(OpenXmlElement rp, string colorCode)
        {
            throw new System.NotImplementedException();
        }

        protected override void SaveSize(OpenXmlElement rp, string sizeValue)
        {
            FontSize fontSize = new FontSize() {Val = sizeValue};
            rp.Append(fontSize);
        }

        protected override void SaveStrike(OpenXmlElement rp)
        {
            Strike strike = new Strike();
            strike.Val = true; 
            rp.Append(strike);
        }

        protected override void SaveColorFill(OpenXmlElement rp, string color)
        {
            Shading shading = new Shading();
            StringValue colorValue = new StringValue() {Value = "auto"};
            StringValue  fillValue = new StringValue() {Value = color};
            shading.Color = colorValue;
            shading.Fill = fillValue;
            shading.Val = ShadingPatternValues.Clear;
            rp.Append(shading);
        }

        public override void RemoveStyles(OpenXmlElement rp)
        {
            RunProperties rpp = (RunProperties) rp;
            rpp?.Bold?.Remove();
            rpp?.FontSize?.Remove();
            rpp?.Strike?.Remove();
            rpp?.Color?.Remove();
            rpp?.Underline?.Remove();
            rpp?.Highlight?.Remove();
            rpp?.Italic?.Remove();
            rpp?.Shading?.Remove();
        }
    }
}
