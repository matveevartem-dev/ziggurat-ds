
using System.Collections.Generic;
using App.Api.Dto;

namespace App.Api.Processor.SaveService
{
    public interface ISaveService<T>
    {
        public void Save(List<List<Segment<T>>> paragraphSegmentList, string filePath, bool isClearStyle);
    }
}
