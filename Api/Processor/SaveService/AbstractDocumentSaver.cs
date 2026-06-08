using System.Collections.Generic;
using System.Linq;
using App.Api.Dto;
using App.Api.Processor.ExtractorService;
using App.Api.Processor.Storage;
using DocumentFormat.OpenXml;

namespace App.Api.Processor.SaveService
{
    public abstract class AbstractDocumentSaver
    {

        protected IOpenXmlFactory _oxmlFactory;
        protected IStyleSaver _styleSaver;
        protected IExtractorService<TextSegmentConfig> _oxmlExtractor;

        abstract protected void RemoveText(OpenXmlCompositeElement run);

        protected void SaveParagraph(
            List<List<Segment<TextSegmentConfig>>> paragraphSegmentList,
            int index,
            IEnumerable<OpenXmlCompositeElement> pList,
            bool isClearStyle
        )
        {
            int paragraphId = paragraphSegmentList.ElementAt(index).First().config.paragraphId;

            RemoveTranslatedContent(pList.ElementAt(paragraphId - 1));

            OpenXmlCompositeElement paragraph = pList.ElementAt(paragraphId - 1);
            OpenXmlCompositeElement lastGeneratedRun = null;
            int endRunId = 0;
            foreach (var segment in paragraphSegmentList.ElementAt(index))
            {
                bool isUseLastRun = endRunId >= segment.config.startRunId;

                int startRunId = segment.config.startRunId;
                endRunId = segment.config.endRunId;

                List<OpenXmlCompositeElement> rList = _oxmlExtractor.GetRunList(paragraph);
                
                if (isClearStyle)
                {
                    foreach (var run in rList)
                    {
                        OpenXmlElement rp = _oxmlExtractor.GetStyleBlock(run);
                        _styleSaver.RemoveStyles(rp);
                    }
                }

                OpenXmlCompositeElement lastSegmentRun = rList.ElementAt(startRunId - 1);
                OpenXmlElement parentRunProperties = _oxmlExtractor.GetStyleBlock(lastSegmentRun);
                OpenXmlCompositeElement useSegmentRun;
                if (
                    lastGeneratedRun != null && 
                    isUseLastRun
                )
                {
                    useSegmentRun = lastGeneratedRun;
                }
                else
                {
                    useSegmentRun = lastSegmentRun;
                }

                lastGeneratedRun = AddStylizedSegment(
                    useSegmentRun, 
                    parentRunProperties,
                    segment
                );

                foreach (var r in rList)
                {
                    RemoveText(r);
                }

                if (isClearStyle)
                {
                    foreach (var run in rList)
                    {
                        OpenXmlElement rp = _oxmlExtractor.GetStyleBlock(run);
                        _styleSaver.RemoveStyles(rp);
                    }
                }
            }
        }

        private OpenXmlCompositeElement AddStylizedSegment(
            OpenXmlCompositeElement firstRun, 
            OpenXmlElement parentRunProperties,
            Segment<TextSegmentConfig> segment
        )
        {
            OpenXmlElement runProperties = parentRunProperties?.CloneNode(true);;

            if (segment.styleList.Any())
            {
                _styleSaver.RemoveStyles(runProperties);
            }

            if (!segment.styleList.Any())
            {
                return CreateRun(
                    segment.content, 
                    runProperties?.CloneNode(true), 
                    firstRun, 
                    segment.content.Length, 
                    0,
                    false
                    );
            }

            OpenXmlCompositeElement intRun = null;
            int endPos = 0;
            bool isTextPartWithoutStyle = false;

            foreach (var item in segment.styleList.Select((value, i) => new { i, value }))
            {
                var style = item.value;
                var i = item.i;

                isTextPartWithoutStyle = style.startPos != endPos; 
                if (isTextPartWithoutStyle)
                {
                    intRun = CreateRun(
                        segment.content, 
                        runProperties?.CloneNode(true), 
                        intRun ?? firstRun, 
                        style.startPos - endPos,
                        endPos,
                        false
                    );
                }
                OpenXmlElement stylizedRunProperties = runProperties?.CloneNode(true);
                if (runProperties == null)
                {
                    stylizedRunProperties = _oxmlFactory.CreateRunProperty();
                }

                foreach (var styleValue in style.styleValueList)
                {
                    _styleSaver.Save(styleValue, stylizedRunProperties);
                }

                intRun = CreateRun(
                    segment.content, 
                    stylizedRunProperties, 
                    intRun ?? firstRun, 
                    style.endPos  -  style.startPos,  
                    style.startPos,
                    false
                );

                endPos = style.endPos;
            }

            Style lastStyle = segment.styleList.Last();
            isTextPartWithoutStyle = segment.content.Length > lastStyle.endPos;

            if (isTextPartWithoutStyle)
            {
                intRun = CreateRun(
                    segment.content, 
                    runProperties?.CloneNode(true), 
                    intRun, 
                    0, 
                    lastStyle.endPos,  
                    false
                    );
            }

            return intRun;
        }

        private OpenXmlCompositeElement CreateRun(
            string text,
            OpenXmlElement rp,
            OpenXmlCompositeElement afterRun,
            int contentLength,
            int styleStartIndex,
            bool addSpace
        )
        {
            rp ??= _oxmlFactory.CreateRunProperty();
            OpenXmlCompositeElement newRun = _oxmlFactory.CreateRun(rp);
            MarkRunAsTranslation(newRun);

            string content;
            if (contentLength == 0 || contentLength + styleStartIndex > text.Length)
            {
                content = text.Substring(styleStartIndex);
            }
            else
            {
                content = text.Substring(styleStartIndex, contentLength);
            }

            OpenXmlElement intText = _oxmlFactory.CreateText(
                    addSpace ? content + " " : content
            );

            newRun.Append(intText);
            afterRun.Parent.InsertAfter(newRun, afterRun);

            return newRun;
        }

        private void RemoveTranslatedContent(OpenXmlCompositeElement paragraph)
        {
            IEnumerable<OpenXmlCompositeElement> runList = paragraph.Descendants<OpenXmlCompositeElement>()
                .Where(
                    r =>
                        r.GetAttributes().FirstOrDefault().Value == ServiceStorage.REPLACED_CRM_CONTENT
                );

            int runListCount = runList.Count() - 1;
            for (var i = runListCount; i >= 0; --i)
            {
                runList.ElementAt(i).Remove();
            }
        }

        private void MarkRunAsTranslation(
            OpenXmlCompositeElement run
        )
        {
            run.SetAttribute(new OpenXmlAttribute(
                ServiceStorage.REPLACED_CRM_CONTENT,
                null,
                ServiceStorage.REPLACED_CRM_CONTENT)
            );
        }
    }
}
