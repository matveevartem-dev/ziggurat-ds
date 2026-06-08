
using App.Api.Processor;
using App.Api.Processor.ExtractorService;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;

namespace App.Api.Dto
{
    public class SegmentListCreator
    {
        private readonly StyleIterator _styleIterator;
        private readonly List<Segment<TextSegmentConfig>> _segmentList = new List<Segment<TextSegmentConfig>>();
        private Segment<TextSegmentConfig> _currentSegment;
        private AbbreviationProcessor _converter;
        private TextSegmentConfig.Location _location;
        private string _pageId;
        private string _content = "";
        private bool _isSegmentCompleted = true;
        private int _segmentId = 1;
        private string _oldStyle;

        public SegmentListCreator(AbbreviationProcessor converter)
        {
            _styleIterator = new StyleIterator();
            _converter = converter;
        }

        public void Setup(TextSegmentConfig.Location location, string pageId = null)
        {
            _pageId = pageId;
            _location = location; 
        }

        public void UpdateCurrentRunId(int runId)
        {
            if (_currentSegment != null)
            {
                if (runId > _currentSegment.config.endRunId && _currentSegment.config.endRunId > 0)
                {
                    runId = _currentSegment.config.endRunId;
                }
                _currentSegment.config.startRunId = runId;    
            }
        }

        public List<Segment<TextSegmentConfig>> GetSegmentList()
        {
            return _segmentList;
        }

        public int GetSegmentId()
        {
            return _segmentId;
        }

        public bool IsSegmentCompleted()
        {
            return _isSegmentCompleted;
        }

        public void MakeCompleteSegment(int paragraphId, int runId, string text = null)
        {
            text ??= _content;
            StartSegment(0, paragraphId, runId);
            CompleteSegment(text, runId);
            _content = "";
        }

        public void MakeSegment(
            string text,
            int paragraphId,
            int runId,
            bool isClosed
        )
        {
            MakeFromTextWithSentenceSeparator(text, paragraphId, runId);

            if (isClosed && _content != "") 
            {
                MakeCompleteSegment(paragraphId, runId);
            }
        }

        private void MakeFromTextWithSentenceSeparator(
            string text,
            int paragraphId,
            int runId
        )
        {
            var _encoded = _converter.Encode(text);

            int charI;
            int segmentStartPos = 0;
            char[] charList = _encoded.ToCharArray();
            char prevChar = _content.LastOrDefault();
            Style multipleSegmentStyle = null;

            for (charI = 0; charI < charList.Count(); charI++)
            {
                char ch = charList[charI];

                if (multipleSegmentStyle != null)
                {
                    multipleSegmentStyle.SetPosition(0, text.Length - charI);
                    _styleIterator.Add(multipleSegmentStyle);
                    multipleSegmentStyle = null;
                }

                if (IsSegmentCompleted())
                {
                    StartSegment(
                        segmentStartPos,
                        paragraphId,
                        runId);
                }

                _content += ch;
                bool isSentenceEnd = IsSentenceEnd(ch, prevChar);
                bool isSeveralSegmentsInRun = isSentenceEnd && 
                                         charI + 1 < text.Length && 
                                         _styleIterator.Last() != null;

                if (isSeveralSegmentsInRun)
                {
                    multipleSegmentStyle = _styleIterator.Last()?.Clone();
                }

                if (isSentenceEnd)
                {
                    _styleIterator.CorrectLastStyle(_content.Length);
                    CompleteSegment(_content, runId);
                    segmentStartPos = charI + 1;
                }

                prevChar = ch;
            }
        }

        public void AddStyle(
            List<StyleValue> styleValueList, 
            int contentLength,
            string innerXml
        )
        {
            bool hasStyle = styleValueList.Count > 0;
            if (!hasStyle)
            {
                ResetOldStyle();
            }

            if (!hasStyle || UpdateLastStyleIfNewTheSame(styleValueList, contentLength))
            {
                return;
            }

            Style style = new Style() {
                segmentId = GetSegmentId(),
                startPos = _content.Length,
                endPos = _content.Length + contentLength,
                value = innerXml,
                styleValueList = styleValueList,
            };
            _styleIterator.Add(style);
        }

        public void ResetOldStyle()
        {
            _oldStyle = null;
        }

        private bool UpdateLastStyleIfNewTheSame(List<StyleValue> styleValueList, int contentLength)
        {
            string serializedStyleList = JsonSerializer.Serialize(styleValueList);
            bool isOldStyleTheSame = serializedStyleList == _oldStyle;
            _oldStyle = serializedStyleList;
            if (isOldStyleTheSame)
            {
                _styleIterator.UpdateLastElement(_content.Length + contentLength);
                return true;
            }
            return false;
        }

        private bool IsSentenceEnd(char ch, char prevChar)
        {
            bool isCharWhitespace = ch == ' ' || ch == Convert.ToChar(160);
            bool prevCharLikeSegmentBreak = IExtractorService<TextSegmentConfig>.AbstractSentenceEndCharList.Contains(prevChar.ToString());
            bool charIsSegmentBreak = IExtractorService<TextSegmentConfig>.SentenceEndCharList.Contains(ch.ToString());

            return charIsSegmentBreak || (
                isCharWhitespace && prevCharLikeSegmentBreak
            );
        }

        private void StartSegment(
            int segmentStartPos,
            int paragraphId, 
            int startRunId
        )
        {
            _isSegmentCompleted = false;

            _currentSegment = new Segment<TextSegmentConfig>()
            {
                id = _segmentId,
                config = new TextSegmentConfig() {
                    segmentStartPos = segmentStartPos,
                    location = _location,
                    paragraphId = paragraphId,
                    startRunId = startRunId,
                    pageId = _pageId
                }
            };
        }

        private void CompleteSegment(
            string content,
            int segmentEndRun
        )
        {
            string _content = _converter.Decode(content);
            //string _content = content;

            _currentSegment.config.endRunId = segmentEndRun;
            _currentSegment.config.segmentLength = content.Length;
            _currentSegment.content = _content;

            if (_currentSegment.config.startRunId > _currentSegment.config.endRunId && _currentSegment.config.endRunId > 0)
            {
                _currentSegment.config.startRunId = _currentSegment.config.endRunId;
            }

            _currentSegment.styleList = _styleIterator.IsEnoughStyles(content.Length) ? 
                _styleIterator.GetStyleList() : 
                new List<Style>();

            if (content != "")
            {
                _segmentList.Add(_currentSegment);
                _segmentId += 1;
            }

            _isSegmentCompleted = true;
            _styleIterator.Clear();
            _content = "";
            _oldStyle = null;
        }
    }
}
