using System.Collections.Generic;

namespace App.Api.Dto
{
    public class StyleValue
    {
        public string value { get; set; }
        public Style.Type type { get; set; }
    }
    
    public class Style
    {
        public enum Type
        {
            Bold, Italic, Highlight, Color, Underline, Size, Strike, ColorFill, LumColor
        }

        public int id { get; set; }
        public int segmentId { get; set; }
        public int startPos { get; set; }
        public int endPos { get; set; }
        public string value { get; set; }
        public List<StyleValue> styleValueList { get; set; } = new List<StyleValue>();

        public Style Clone()
        {
            return (Style)this.MemberwiseClone();
        }

        public void SetPosition(int start, int end)
        {
            startPos = start;
            endPos = end;
        }
    }
}
