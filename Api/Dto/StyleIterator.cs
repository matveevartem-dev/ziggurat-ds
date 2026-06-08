using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace App.Api.Dto
{
    public class StyleIterator
    {
        private List<Style> _styleList = new List<Style>();
        private int _id = 0;

        public void Clear()
        {
            _styleList = new List<Style>();
            _id = 0;
        }
        public List<Style> GetStyleList()
        {
            return _styleList;
        }
        public void Add(
            Style style
            )
        {
            if (null == style)
            {
                return;
            }
            ++_id;
            style.id = _id;
            _styleList.Add(style);
        }
        public bool IsNotEmpty()
        {
            return _styleList.Count > 0;
        }

        public Style Last()
        {
            return _styleList.LastOrDefault();
        }

        public bool IsEnoughStyles(int contentLength)
        {
            if (_styleList.Count == 1)
            {
                Style style = _styleList.First();
                int styleLength = style.endPos - style.startPos;
                if (contentLength > styleLength)
                {
                    return true;
                }
            }

            bool isEnoughStyles = _styleList.Count > 1;

            return isEnoughStyles;
        }
        public void UpdateLastElement(int endPos)
        {
            if (false == IsNotEmpty())
            {
                return;
            }
            Last().endPos = endPos;
        }

        public string GetLastStyleAsJson()
        {
            Style style = _styleList.LastOrDefault();
            if (style != null)
            {
                return JsonSerializer.Serialize(style.styleValueList);
            }

            return null;
        }

        public void CorrectLastStyle(int endPos)
        {
            Style style = Last();
            if (style == null)
            {
                return;
            }

            style.endPos = endPos;
        }
    }
}
