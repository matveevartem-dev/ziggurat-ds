using System.Collections.Generic;
using System.Linq;
using App.Api.Dto;
using App.Api.Processor.Storage;
using DocumentFormat.OpenXml;

namespace App.Api.Processor.ExtractorService
{
    public abstract class AbstractDocumentExtractor<
        THyperlink, 
        TText, 
        TNumberingProperties,
        TRun
    > : IExtractorService<TextSegmentConfig>
    {
        protected IStyleExtractor OxmlStyleExtractor;
        protected SegmentListCreator SegmentListCreator;

        public abstract List<Segment<TextSegmentConfig>> Extract(string filePath);

        /// <summary>
        /// Определяет наличие табуляции в секции Run
        /// </summary>
        /// <param name="r"></param>
        /// <returns></returns>
        protected abstract bool RunHasTabChar(OpenXmlCompositeElement r);

        public void MakeSegmentList(IEnumerable<OpenXmlCompositeElement> pList)
        {
            foreach (var p in pList.Select((value, i) => new { i, value }))
            {
                bool paragraphHasNumbering = p.value
                    .Elements()
                    .FirstOrDefault(n => n is TNumberingProperties) != null;

                List<OpenXmlCompositeElement> rList = GetRunList(p.value);
                bool isTextMustClear = true;
                string innerText = "";
                foreach (var r in rList.Select((value, i) => new {i, value}))
                {
                    OpenXmlElement rp = OxmlStyleExtractor.GetStyleBlock(r.value);
                    List<OpenXmlElement> tList = r.value.Elements().Where(e => e is TText).ToList();
                    innerText = isTextMustClear ? "" : innerText;
                    foreach (var t in tList)
                    {
                        innerText += t.InnerText;
                    }

                    if (paragraphHasNumbering)
                    {
                        innerText = innerText + "  ";
                        paragraphHasNumbering = false;
                    }

                    OpenXmlCompositeElement nextRun = rList.ElementAtOrDefault(r.i + 1);
                    bool nextRunHasTabChar = RunHasTabChar(nextRun);

                    // innerText = nextRunHasTabChar ? innerText + ' ' : innerText;

                    if (rp != null && innerText != "")
                    {
                        SegmentListCreator.AddStyle(
                            OxmlStyleExtractor.Extract(rp),
                            innerText.Length,
                            rp.InnerXml
                        );
                    }
                    else
                    {
                        SegmentListCreator.ResetOldStyle();
                    }

                    if (innerText.Trim() == "")
                    {
                        SegmentListCreator.UpdateCurrentRunId(r.i + 2);
                    }

                    if (r.value.Parent is THyperlink)
                    {
                        isTextMustClear = false;
                        if (nextRun?.Parent != r.value.Parent || nextRunHasTabChar)
                        {
                            SegmentListCreator.MakeCompleteSegment(
                                p.i + 1, 
                                r.i + 1, 
                                innerText
                                );
                            isTextMustClear = true;
                        }

                        continue;
                    }
                    
                    bool nextRunIsHyperlink = nextRun?.Parent is THyperlink;
                    bool isLastRun = nextRun == null || (nextRunIsHyperlink || nextRunHasTabChar);
                    SegmentListCreator.MakeSegment(
                        innerText, 
                        p.i + 1, 
                        r.i + 1, 
                        isLastRun
                        );
                    isTextMustClear = true;
                }
            }
        }

        public List<OpenXmlCompositeElement> GetRunList(
            OpenXmlCompositeElement p
        )
        {
            List<OpenXmlCompositeElement> runList = p.Descendants<OpenXmlCompositeElement>().Where(
                r =>
                    (
                    string.IsNullOrEmpty(r.GetAttributes().FirstOrDefault().Value) ||
                    r.GetAttributes().FirstOrDefault().Value != ServiceStorage.REPLACED_CRM_CONTENT
                    //r.GetAttributes().FirstOrDefault() == null || 
                    //r.GetAttributes().FirstOrDefault() != null && 
                    //r.GetAttributes().FirstOrDefault().Value != ServiceStorage.REPLACED_CRM_CONTENT
                    ) &&
                    r is TRun &&
                    r.Ancestors<AlternateContentFallback>().FirstOrDefault() == null &&
                    r.Descendants<AlternateContentFallback>().FirstOrDefault() == null &&
                    (
                        r.Parent == p ||
                        r.Parent is THyperlink
                    )
            ).ToList();

            return runList;
        }

        public OpenXmlElement GetParentHyperlink(OpenXmlCompositeElement run)
        {
            if (run.Parent is THyperlink)
            {
                return run.Parent;
            }

            return null;
        }

        public int CalculateContentLength(List<Segment<TextSegmentConfig>> segmentList)
        {
            int contentLength = 0;
            foreach (Segment<TextSegmentConfig> segment in segmentList)
            {
                contentLength = contentLength + segment.content.Length;
            }

            return contentLength;
        }

        public OpenXmlElement GetStyleBlock(OpenXmlElement run)
        {
            return OxmlStyleExtractor.GetStyleBlock(run);
        }
    }
}
