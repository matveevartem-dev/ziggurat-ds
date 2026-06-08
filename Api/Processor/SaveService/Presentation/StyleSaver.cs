using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using Highlight = DocumentFormat.OpenXml.Drawing.Highlight;
using RunProperties = DocumentFormat.OpenXml.Drawing.RunProperties;

namespace App.Api.Processor.SaveService.Presentation
{
    public class StyleSaver: AbstractStyleSaver
    {
        protected override void SaveBold(OpenXmlElement rp)
        {
            RunProperties rpp = (RunProperties) rp;
            rpp.Bold = true;
        }

        protected override void SaveItalic(OpenXmlElement rp)
        {
            RunProperties rpp = (RunProperties) rp;
            rpp.Italic = true;
        }

        protected override void SaveHighlight(OpenXmlElement rp, string color)
        {
            RunProperties rpp = (RunProperties) rp;
            Highlight highlight = new Highlight();
            RgbColorModelHex colorModelHex = new RgbColorModelHex();
            colorModelHex.Val = color;
            
            highlight.RgbColorModelHex = colorModelHex;
            rpp.InsertAt(highlight, 0);
        }

        protected override void SaveUnderline(OpenXmlElement rp, string style)
        {
            RunProperties rpp = (RunProperties) rp;
            rpp.Underline = TextUnderlineValues.Single;
        }

        protected override void SaveColor(OpenXmlElement rp, string colorCode)
        {
            RunProperties rpp = (RunProperties) rp;
            SolidFill color = new SolidFill() {RgbColorModelHex = new RgbColorModelHex() {Val = colorCode}};
            
            rpp.InsertAt(color, 0);                
        }

        protected override void SaveLumColor(OpenXmlElement rp, string colorCode)
        {
            RunProperties rpp = (RunProperties) rp;
            SchemeColor schemeColor = new SchemeColor() {Val = SchemeColorValues.Accent1};
            
            schemeColor.AppendChild(new LuminanceModulation() {Val = Int32.Parse(colorCode)});
            
            SolidFill color = new SolidFill() {SchemeColor = schemeColor};
            
            rpp.InsertAt(color, 0);       
        }

        protected override void SaveSize(OpenXmlElement rp, string size)
        {
            RunProperties rpp = (RunProperties) rp;
            Int32Value fontSize = new Int32Value(int.Parse(size));
            rpp.FontSize = fontSize;    
        }

        protected override void SaveStrike(OpenXmlElement rp)
        {
            RunProperties rpp = (RunProperties) rp;
            rpp.Strike = TextStrikeValues.SingleStrike;
        }

        protected override void SaveColorFill(OpenXmlElement rp, string color)
        {
            return; 
        }

        public override void RemoveStyles(OpenXmlElement rp)
        {
            RunProperties rpp = (RunProperties) rp;
            rpp.Bold = false;
            rpp.Strike = null;
            rpp.Underline = null;
            rpp.Italic = null;

            SolidFill color = rpp.Elements<SolidFill>().FirstOrDefault();
            color?.Remove();
            
            Highlight highlight = rpp.Elements<Highlight>().FirstOrDefault();
            highlight?.Remove();
        }
    }
}
