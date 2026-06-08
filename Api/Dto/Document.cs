using System;
using System.Collections.Generic;

namespace App.Api.Dto
{
    [Serializable]
    public class Document<T> : IDocument
    {
        public enum DocumentType
        {
            Text, Spreadsheet 
        }

        public DocumentType type { get; set; }
        public IEnumerable<Segment<T>> segmentList { get; set; }
        public int charCount { get; set; }
    }
}
