namespace App.Api.Dto
{
    public class TextSegmentConfig
    {
        public enum Location
        {
            Document, Footnotes, Endnotes, Footer, Header
        }
        public int paragraphId { get; set; }
        public string pageId { get; set; }
        public int startRunId { get; set; }
        public int endRunId { get; set; }
        public int textId { get; set; }
        public Location location { get; set; }
        public int segmentLength { get; set; }
        public int segmentStartPos { get; set; }
    }
}
