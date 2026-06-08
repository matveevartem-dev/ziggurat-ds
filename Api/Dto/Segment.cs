using System.Collections.Generic;

namespace App.Api.Dto
{
    public class Segment<T>
    {
        public int id { get; set; }
        public string content { get; set; }
        public IEnumerable<Style> styleList { get; set; } = new List<Style>();
        public T config { get; set; }
        public bool hasLineBreak { get; set; }
    }
}
