using System;
using App.Api.Dto;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Style = App.Api.Dto.Style;

namespace App.Api.Processor.SaveService
{
    public abstract class AbstractStyleSaver: IStyleSaver
    {
        protected abstract void SaveBold(OpenXmlElement rp);
        protected abstract void SaveItalic(OpenXmlElement rp);
        protected abstract void SaveHighlight(OpenXmlElement rp, string color);
        protected abstract void SaveUnderline(OpenXmlElement rp, string style);
        protected abstract void SaveColor(OpenXmlElement rp, string colorCode);
        protected abstract void SaveLumColor(OpenXmlElement rp, string colorCode);
        protected abstract void SaveSize(OpenXmlElement rp, string sizeValue);
        protected abstract void SaveStrike(OpenXmlElement rp);
        protected abstract void SaveColorFill(OpenXmlElement rp, string color);

        public abstract void RemoveStyles(OpenXmlElement rp);

        public void Save(StyleValue styleValue, OpenXmlElement rp)
        {
            if (styleValue.type == Style.Type.Bold)
            {
                SaveBold(rp);
            }
            if (styleValue.type == Style.Type.Italic)
            {
                SaveItalic(rp);
            }
            if (styleValue.type == Style.Type.Highlight)
            {
                SaveHighlight(rp, styleValue.value);
            }
            if (styleValue.type == Style.Type.Color)
            {
                SaveColor(rp, styleValue.value);
            }
            if (styleValue.type == Style.Type.Underline)
            {
                SaveUnderline(rp, styleValue.value);
            }
            if (styleValue.type == Style.Type.Strike)
            {
                SaveStrike(rp);
            }
            if (styleValue.type == Style.Type.Size)
            {
                SaveSize(rp, styleValue.value);
            }
            if (styleValue.type == Style.Type.ColorFill)
            {
                SaveColorFill(rp, styleValue.value);
            }

            if (styleValue.type == Style.Type.LumColor)
            {
                SaveLumColor(rp,styleValue.value);
            }
        }
    }
}
