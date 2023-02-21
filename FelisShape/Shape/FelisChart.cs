using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using System.Text.RegularExpressions;
using System.Xml;
using System.Diagnostics;
using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;
using FelisOpenXml.FelisShape.Base;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Runtime.ConstrainedExecution;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The chart shape
    /// </summary>
    [FelisGraphicFrame(@"http://schemas.openxmlformats.org/drawingml/2006/chart")]
    public class FelisChart : FelisGraphicFrame
    {
        /// <summary>
        /// The part of the chart
        /// </summary>
        public readonly ChartPart ChartPart;
        /// <summary>
        /// The space element of the root of the part
        /// </summary>
        public readonly C.ChartSpace ChartSpace;
        /// <summary>
        /// The element of the chart
        /// </summary>
        public readonly OpenXmlCompositeElement ChartElement;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_element">The graphic element using the chart</param>
        /// <exception cref="OpenXmlPackageException"></exception>
        protected FelisChart(P.GraphicFrame _element)
            : base(_element)
        {
            var slidePart = FelisSlide.RetrospectToSlideElement(_element)?.SlidePart;
            var chartRef = ForceGraphicData.GetFirstChild<C.ChartReference>() ?? ForceGraphicData.InsertElement(new C.ChartReference());
            if (null == slidePart)
            {
                throw new OpenXmlPackageException("Can not locate the part of the slide.");
            }
            else if (slidePart.TryGetPartById(chartRef.Id?.Value ?? string.Empty, out OpenXmlPart? part) && (part is ChartPart chartPart))
            {
                ChartPart = chartPart;
            }
            else
            {
                ChartPart = slidePart.AddNewPart<ChartPart>();
                ChartPart.ChartSpace = new C.ChartSpace();
                chartRef.Id = slidePart!.GetIdOfPart(ChartPart);
            }
            ChartSpace = ChartPart.ChartSpace;
            ChartElement = ForceGetChartElement(ChartSpace);
            seriesCollection = new FelisChartSeriesCollection(ChartElement);
        }

        /// <summary>
        /// Get the chart element in the chart
        /// A new bar chart will be created if there is no existed chart element
        /// </summary>
        /// <param name="_chartSpace">The chart space element which contains the chart element</param>
        /// <returns></returns>
        protected static OpenXmlCompositeElement ForceGetChartElement(C.ChartSpace _chartSpace)
        {
            var chart = _chartSpace.GetFirstChild<C.Chart>() ?? _chartSpace.InsertElement(new C.Chart());
            var plotArea = chart.PlotArea ?? (chart.PlotArea = new C.PlotArea());
            var chartElement = plotArea.ChildElements.FirstOrDefault(e => e.LocalName.EndsWith("Chart", StringComparison.Ordinal) && (e is OpenXmlCompositeElement) && e.NamespaceUri == chart.NamespaceUri);
            if (null == chartElement)
            {
                plotArea.AddChild(chartElement = new C.BarChart(), false);
            }
            return (chartElement as OpenXmlCompositeElement)!;
        }

        /// <summary>
        /// The cache of the collection of the series.
        /// </summary>
        protected readonly FelisChartSeriesCollection seriesCollection;

        /// <summary>
        /// The collection of the series in the chart
        /// </summary>
        public FelisChartSeriesCollection SeriesCollection => seriesCollection;

        /// <summary>
        /// The categories of the chart.
        /// This value is taken as the categories declared in the first series of the chart
        /// </summary>
        public FelisChartCategories? PrimaryCategories => SeriesCollection[0]?.Categories;

        /// <summary>
        /// The title of the chart
        /// </summary>
        public string? Title
        {
            get
            {
                C.Title? titleElement = ChartPart.ChartSpace?.GetFirstChild<C.Chart>()?.Title;
                if (null != titleElement)
                {
                    C.ChartText? chartTextElement = titleElement.ChartText;
                    if (null != chartTextElement)
                    {
                        // Static title
                        string? title = chartTextElement.RichText?.Descendants<A.Text>().Select(t => t.Text)
                            .Aggregate((t1, t2) => t1 + t2);
                        if (null != title)
                        {
                            return title;
                        }

                        // Dynamic title
                        return chartTextElement.Descendants<C.StringPoint>().Single().InnerText;
                    }

                    // PieChart uses only one series for view.
                    // However, it can have store multiple series data in the spreadsheet.
                    if (ChartElement is C.PieChart)
                    {
                        return SeriesCollection[0]?.Title;
                    }
                }

                return null;
            }

            set
            {
                if (null != value)
                {
                    var titleElement = ChartPart!.ChartSpace?.GetFirstChild<C.Chart>()?.Title;
                    if (null != titleElement)
                    {
                        var chartText = titleElement.ChartText;
                        var richText = chartText?.RichText;
                        if (null == richText)
                        {
                            if (null == chartText)
                            {
                                chartText = new C.ChartText();
                                titleElement.AddChild(chartText, false);
                            }
                            richText = new C.RichText(
                                new A.BodyProperties(),
                                new A.ListStyle(),
                                new A.Paragraph(
                                    new A.ParagraphProperties(),
                                    new A.Run(
                                        new A.Text(value)
                                    ),
                                    new A.EndParagraphRunProperties()
                                )
                            );
                            chartText.AddChild(richText, false);
                        }
                        else
                        {
                            var paragraphs = richText.Descendants<A.Paragraph>().ToArray();
                            int index = 0;
                            foreach (var text in value!.Split("\n"))
                            {
                                var paragraph = paragraphs.ElementAtOrDefault(index++);
                                if (null != paragraph)
                                {
                                    var runs = paragraph.Descendants<A.Run>().ToArray();
                                    bool firstRun = true;
                                    foreach (var run in runs)
                                    {
                                        if (firstRun)
                                        {
                                            if (null != run.Text)
                                            {
                                                run.Text.Text = text;
                                            }
                                            firstRun = false;
                                        }
                                        else
                                        {
                                            run.Remove();
                                        }
                                    }
                                    if (firstRun)
                                    {
                                        paragraph.AddChild(new A.Run(
                                                new A.Text(value)
                                            ), false
                                        );
                                    }
                                }
                                else
                                {
                                    richText.AddChild(
                                        new A.Paragraph(
                                            new A.Run(
                                                new A.Text(value)
                                            )
                                        ), false
                                    );
                                }
                            }
                            for (; index < paragraphs.Length; index++)
                            {
                                paragraphs[index].Remove();
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Remove the external workbook
        /// </summary>
        public void RemoveExternalWorkbook()
        {
            var extData = ChartSpace.GetFirstChild<C.ExternalData>();
            if (null != extData) 
            {
                try
                {
                    if (ChartPart.TryGetPartById(extData.Id?.Value ?? String.Empty, out OpenXmlPart? part))
                    {
                        ChartPart.DeletePart(part);
                    }
                }
                catch (Exception err)
                {
                    Trace.TraceWarning(err.ToString());
                }
                finally
                {
                    extData.Remove();
                }
            }
        }

        /// <summary>
        /// Set the external workbook
        /// </summary>
        /// <param name="_source"></param>
        /// <param name="_autoUpdate"></param>
        public void SetExternalWorkBook(Stream _source, bool _autoUpdate = false)
        {
            var extPart = ChartPart.AddNewPart<EmbeddedPackagePart>("http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", null!);
            extPart.FeedData(_source);
            var rid = ChartPart.GetIdOfPart(extPart);

            var extData = ChartSpace.ForceGetChild<C.ExternalData>();
            if (null != extData)
            {
                extData.Id = rid;
                var updateNode = extData.ForceGetChild<C.AutoUpdate>();
                if (null != updateNode)
                {
                    updateNode.Val = _autoUpdate;
                }
            }
        }

        /// <summary>
        /// Work with the content of the external workbook
        /// </summary>
        /// <param name="_fn"></param>
        public void WorkWithExternalWorkbook(Action<Stream> _fn)
        {
            var extData = ChartSpace.GetFirstChild<C.ExternalData>();
            if (null != extData)
            {
                try
                {
                    if (ChartPart.TryGetPartById(extData.Id?.Value ?? String.Empty, out OpenXmlPart? part))
                    {
                        using (var stream = part.GetStream())
                        {
                            _fn(stream);
                        }
                    }
                }
                catch (Exception err)
                {
                    Trace.TraceWarning(err.ToString());
                }
            }
        }

        /// <summary>
        /// Submit the changeing of the chart
        /// </summary>
        public void Submit()
        {
            ChartSpace.Save();
        }
    }

    /// <summary>
    /// The class for manipulating the series of the data
    /// </summary>
    public class FelisChartSeries : FelisCompositeElement
    {
        private Lazy<FelisChartCategories> categories;
        private Lazy<FelisChartValues> values;

        internal FelisChartSeries(OpenXmlCompositeElement _element)
            : base(_element) 
        {
            categories = new Lazy<FelisChartCategories>(() =>
            {
                var ret = Element.GetFirstChild<C.CategoryAxisData>();
                if (null == ret)
                {
                    Element.AddChild(ret = new C.CategoryAxisData(), false);
                }
                return new FelisChartCategories(ret);
            });
            values = new Lazy<FelisChartValues>(() =>
            {
                var ret = Element.GetFirstChild<C.Values>();
                if (null == ret)
                {
                    Element.AddChild(ret = new C.Values(), false);
                }
                return new FelisChartValues(ret);
            });
        }

        /// <summary>
        /// The title of the series
        /// </summary>
        public string? Title
        {
            get
            {
                return Element.GetFirstChild<C.SeriesText>()?.StringReference?.StringCache?.GetFirstChild<C.StringPoint>()?.NumericValue?.Text;
            }
            set
            {
                var txtValue = Element.ForceGetChild<C.SeriesText>()!.ForceGetChild<C.StringReference>()?.ForceGetChild<C.StringCache>()?.ForceGetChild<C.StringPoint>()?.NumericValue;
                if (null != txtValue)
                {
                    txtValue.Text = value ?? string.Empty;
                }
            }
        }

        /// <summary>
        /// Set the data reference of the series' title
        /// </summary>
        /// <param name="_book"></param>
        /// <param name="_sheet"></param>
        /// <param name="_cell"></param>
        public void SetTitleReference(string? _book, string? _sheet, string _cell)
        {
            var refStr = _cell.Trim();
            if (!string.IsNullOrEmpty(_cell))
            {
                var strRef = Element.GetFirstChild<C.SeriesText>()?.StringReference;
                var formula = strRef?.ForceGetChild<C.Formula>(() => strRef.InsertElement(new C.Formula(), 0));
                if (null != formula)
                {
                    if (!string.IsNullOrWhiteSpace(_sheet))
                    {
                        refStr = string.IsNullOrWhiteSpace(_book) ? $"{_sheet.Trim()}!{refStr}" : $"\'[{_book.Trim()}]{_sheet.Trim()}\'!{refStr}";
                    }
                    formula.Text = refStr;
                }
            }
        }

        /// <summary>
        /// Remove the reference setting of the title
        /// </summary>
        public void ClearTitleReference()
        {
            var formula = Element.GetFirstChild<C.SeriesText>()?.StringReference?.GetFirstChild<C.Formula>();
            formula?.Remove();
        }

        /// <summary>
        /// The index of the series in the chart
        /// </summary>
        public uint Index
        {
            get
            {
                var index = Element.GetFirstChild<C.Index>();
                return (null != index) ? (index.Val ?? 0 + 1) : 0;
            }
        }

        /// <summary>
        /// Get the categories of this series
        /// </summary>
        public FelisChartCategories Categories => categories.Value;

        /// <summary>
        /// Get the values of this series
        /// </summary>
        public FelisChartValues Values => values.Value;

        /// <summary>
        /// Remove the reference information of the series
        /// </summary>
        public void RemoveDataReference()
        {
            Categories.RemoveReference();
            Values.RemoveReference();
            Element.GetFirstChild<C.SeriesText>()?.StringReference?.RemoveAllChildren<C.Formula>();
        }
    }

    /// <summary>
    /// The collection of the series in the chart
    /// </summary>
    public class FelisChartSeriesCollection : FelisModifiableCollection<OpenXmlCompositeElement, OpenXmlCompositeElement, FelisChartSeries>
    {
        private IReadOnlyList<OpenXmlCompositeElement?>? templateNodes = null;
        private Type? placeSiblingNodeType = null;

        internal FelisChartSeriesCollection(OpenXmlCompositeElement _container) : base(_container)
        {
        }

        /// <summary>
        /// Boxing the series element to series object
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override FelisChartSeries BoxingElement(OpenXmlCompositeElement _element)
        {
            return new FelisChartSeries(_element);
        }

        /// <summary>
        /// Create a new series element
        /// </summary>
        /// <param name="_index"></param>
        /// <returns></returns>
        /// <exception cref="Exception">The template is error</exception>
        protected override OpenXmlCompositeElement CreateElement(int _index)
        {
            if ((null == templateNodes) || (templateNodes.Count <= 0))
            {
                throw new Exception("Template is empty");
            }
            int templateIdx = _index % templateNodes.Count;
            var newNode = templateNodes[templateIdx]?.CloneNode(true) as OpenXmlCompositeElement;
            if (null == newNode)
            {
                throw new Exception("Can not clone from the template");
            }
            var indexElement = newNode.GetFirstChild<C.Index>();
            if (null != indexElement)
            {
                indexElement.Val = (uint)_index;
            }
            var orderElement = newNode.GetFirstChild<C.Order>();
            if (null != orderElement)
            {
                orderElement.Val = (uint)_index;
            }
            return newNode;
        }

        /// <summary>
        /// Unboxing the series object to series element
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override OpenXmlElement UnboxingElement(FelisChartSeries _element)
        {
            return _element.Element;
        }

        /// <summary>
        /// Get an iterator of the series in the chart
        /// </summary>
        /// <returns></returns>
        protected override IEnumerable<OpenXmlCompositeElement> GetElements()
        {
            List<OpenXmlCompositeElement?>? templates = ((null == templateNodes) ? new List<OpenXmlCompositeElement?>() : null);
            OpenXmlElement? lastSibling = null;
            if (null != templates)
            {
                templateNodes = templates;
            }
            foreach (var item in ContainerElement.Elements())
            {
                if ((item is OpenXmlCompositeElement serElement) && item.LocalName.Equals("ser", StringComparison.Ordinal))
                {
                    if (null != templates)
                    {
                        templates.Add(serElement.CloneNode(true) as OpenXmlCompositeElement);
                        lastSibling = serElement.NextSibling();
                    }
                    yield return serElement;
                }
            }

            if (null != templates)
            {
                placeSiblingNodeType = lastSibling?.GetType();
            }
        }

        /// <summary>
        /// Add the first element into the container
        /// </summary>
        /// <param name="_newElement"></param>
        protected override void PlaceFirstElement(OpenXmlElement _newElement)
        {
            if (null == placeSiblingNodeType)
            {
                ContainerElement.AddChild(_newElement, false);
            }
            else
            {
                var node = ContainerElement.GetFirstChild(e => placeSiblingNodeType.IsInstanceOfType(e));
                if (null == node)
                {
                    ContainerElement.AddChild(_newElement, false);
                }
                else
                {
                    ContainerElement.InsertBefore(_newElement, node);
                }
            }
        }

        /// <summary>
        /// Reset the index of the series after some one is add or delete from the collection
        /// </summary>
        /// <param name="_list"></param>
        /// <param name="_startIndex"></param>
        protected override void ResetIndexOfElements(IEnumerable<FelisChartSeries> _list, int _startIndex)
        {
            uint index = (uint)_startIndex;
            foreach (var ser in _list)
            {
                var indexElement = ser.Element.GetFirstChild<C.Index>();
                if (null != indexElement)
                {
                    indexElement.Val = index;
                }
                var orderElement = ser.Element.GetFirstChild<C.Order>();
                if (null != orderElement)
                {
                    orderElement.Val = index;
                }
                index++;
            }
        }
    }

    /// <summary>
    /// The basic class for manipulating the element of the reference, such as C.DataReference, C.StringReference
    /// </summary>
    public abstract class FelisChartDataReferance
    {
        /// <summary>
        /// The SDK element of the reference
        /// </summary>
        public readonly OpenXmlCompositeElement ReferenceElement;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_refElement"></param>
        protected FelisChartDataReferance(OpenXmlCompositeElement _refElement)
        {
            ReferenceElement = _refElement;
            ReloadCache();
        }

        #region 倒置依赖声明
        /// <summary>
        /// Check if the input element is the typeof the data cache
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected abstract bool CheckIsCache(OpenXmlElement _element);

        /// <summary>
        /// Check if the input element is the typeof the data point
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected abstract bool CheckIsPointer(OpenXmlElement _element);

        /// <summary>
        /// Create a new data point element
        /// </summary>
        /// <returns></returns>
        protected abstract OpenXmlElement CreatePointer();
        #endregion

        #region 值处理
        /// <summary>
        /// The cache of the values
        /// </summary>
        private Lazy<string?[]>? valuesCache;

        /// <summary>
        /// The loader for the values
        /// </summary>
        /// <returns></returns>
        private string?[] ValuesCacheLoader()
        {
            return ReferenceElement.GetFirstChild(CheckIsCache)?.Children(CheckIsPointer)
                        .OrderBy(e => GetPointIndex(e))
                        .Select(e => e.GetFirstChild<C.NumericValue>()?.Text)
                        .ToArray() ?? Array.Empty<string>();
        }

        /// <summary>
        /// Reload the cache of the values
        /// </summary>
        public void ReloadCache()
        {
            Interlocked.Exchange(ref valuesCache, new Lazy<string?[]>(ValuesCacheLoader));
        }

        /// <summary>
        /// Get the index value in the data point
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected static uint GetPointIndex(OpenXmlElement _element)
        {
            return ((_element is C.StringPoint strPt)
                        ? strPt.Index?.Value
                        : (_element is C.NumericPoint numPt) ? numPt.Index?.Value : uint.MaxValue)
                   ?? uint.MaxValue;
        }

        /// <summary>
        /// Set the index of a data point
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_index"></param>
        protected static void SetPointIndex(OpenXmlElement _element, uint _index)
        {
            if (_element is C.StringPoint strPt)
            {
                strPt.Index = _index;
            }
            else if (_element is C.NumericPoint numPt)
            {
                numPt.Index = _index;
            }
        }

        /// <summary>
        /// Get the values
        /// </summary>
        public object?[] Values => valuesCache?.Value ?? Array.Empty<string?>();

        /// <summary>
        /// Get the count of the values
        /// </summary>
        public int ValuesCount => valuesCache?.Value.Length ?? 0;

        /// <summary>
        /// Get or set the value in the special index
        /// </summary>
        /// <param name="_index"></param>
        /// <returns></returns>
        public object? this[int _index]
        {
            get
            {
                return _index >= 0 && _index < valuesCache?.Value.Length ? valuesCache.Value[_index] : null;
            }
            set
            {
                if (_index >= 0 && _index < valuesCache?.Value.Length)
                {
                    string data = null != value ? value!.ToString()! : string.Empty;
                    valuesCache!.Value[_index] = data;
                    var valElement = ReferenceElement.GetFirstChild(CheckIsCache)?.GetFirstChild(e => CheckIsPointer(e) && (GetPointIndex(e) == _index))?.GetFirstChild<C.NumericValue>();
                    if (valElement != null)
                    {
                        valElement.Text = data;
                    }
                }
            }
        }

        /// <summary>
        /// Rewrite all the values. 
        /// This method can change the count of the values
        /// </summary>
        /// <param name="_values"></param>
        /// <returns></returns>
        public virtual OpenXmlElement? ReWrite(IEnumerable<object?> _values)
        {
            var cache = ReferenceElement.GetFirstChild(CheckIsCache);
            if (null != cache)
            {
                cache.RemoveAllChildren<C.PointCount>();
                cache.RemoveAllChildren(CheckIsPointer);
                uint index = 0;
                OpenXmlElement? firstPt = null;
                foreach (var item in _values)
                {
                    string data = null == item ? string.Empty : item.ToString()!;
                    var ptr = CreatePointer();
                    if (null == firstPt)
                    {
                        firstPt = ptr;
                    }
                    SetPointIndex(ptr, index++);
                    ptr.InsertAt(new C.NumericValue(data), 0);
                    cache.AppendChild(ptr);
                }
                if (null != firstPt)
                {
                    cache.InsertBefore(new C.PointCount() { Val = index }, firstPt);
                }
                else
                {
                    cache.InsertElement(new C.PointCount() { Val = index }, 0);
                }
            }

            ReloadCache();
            return cache;
        }

        /// <summary>
        /// Rewrite all the values. 
        /// This method can change the count of the values
        /// </summary>
        /// <param name="_values"></param>
        public void ReWrite(params object?[] _values)
        {
            ReWrite(_values as IEnumerable<object?>);
        }
        #endregion

        #region 格式处理
        /// <summary>
        /// The default format of the value
        /// </summary>
        public static readonly string DefaultFormat = "General";

        /// <summary>
        /// Get or set the format of the value
        /// </summary>
        public string? Format
        {
            get
            {
                return ReferenceElement.GetFirstChild(CheckIsCache)?.GetFirstChild<C.FormatCode>()?.Text;
            }
            set
            {
                if (!string.IsNullOrWhiteSpace(value))
                {
                    var cache = ReferenceElement.GetFirstChild(CheckIsCache);
                    if ((cache is C.NumberingCache) || (cache is C.NumberLiteral))
                    {
                        var fmtCode = cache.GetFirstChild<C.FormatCode>();
                        if (null == fmtCode)
                        {
                            fmtCode = cache.InsertElement(new C.FormatCode(), 0);
                        }
                        if (null != fmtCode)
                        {
                            fmtCode.Text = value;
                        }
                    }
                }
            }
        }
        #endregion

        #region 关于引用位置的处理
        /// <summary>
        /// Get or set the reference of the data
        /// </summary>
        public (string? book, string? sheet, string start, string? end) DataReference
        {
            get
            {
                var formula = ReferenceElement.GetFirstChild<C.Formula>();
                if (null != formula)
                {
                    var match = Regex.Match(formula.Text, @"'?(?:\[(.+)\])?([^!]+)'?!([^:]+)(?::([^:]+))?");
                    return (book: match.Groups[1].Value, sheet: match.Groups[2].Value, start: match.Groups[3].Value, end: match.Groups[4].Value);
                }
                else
                {
                    return (book: null, sheet: string.Empty, start: string.Empty, end: string.Empty);
                }
            }
            set
            {
                string refStr = (string.IsNullOrWhiteSpace(value.end) ? value.start : $"{value.start}:{value.end}").Trim();
                if (!string.IsNullOrEmpty(refStr))
                {
                    var formula = ReferenceElement.ForceGetChild<C.Formula>(() => ReferenceElement.InsertElement(new C.Formula(), 0));
                    if (null != formula)
                    {
                        if (!string.IsNullOrWhiteSpace(value.sheet))
                        {
                            refStr = string.IsNullOrWhiteSpace(value.book) ? $"{value.sheet}!{refStr}" : $"\'[{value.book}]{value.sheet}\'!{refStr}";
                        }
                        formula.Text = refStr;
                    }
                }
            }
        }

        /// <summary>
        /// Remove the reference of the data
        /// </summary>
        public void RemoveReference()
        {
            ReferenceElement.RemoveAllChildren<C.Formula>();
        }

        /// <summary>
        /// Set the reference to unknown
        /// </summary>
        /// <param name="_start">The loaction of the start data</param>
        /// <param name="_end">The location of the end data</param>
        public void SetUnknownReference(string _start, string _end)
        {
            DataReference = (book: "unknown", sheet: "unknown", start: _start, end: _end);
        }
        #endregion
    }

    /// <summary>
    /// The class for manipulating the values of the numeric data
    /// </summary>
    public class FelisChartNumberReference : FelisChartDataReferance
    {
        internal FelisChartNumberReference(C.NumberReference _element)
            : base(_element)
        {
        }

        /// <summary>
        /// Rewrite all the values
        /// </summary>
        /// <param name="_values"></param>
        public override OpenXmlElement? ReWrite(IEnumerable<object?> _values)
        {
            var cache = base.ReWrite(_values);
            if (null != cache)
            {
                if (!cache.Elements<C.FormatCode>().Any())
                {
                    cache.InsertElement(new C.FormatCode(DefaultFormat), 0);
                }
            }
            return cache;
        }

        /// <summary>
        /// Check if the element is the cache
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override bool CheckIsCache(OpenXmlElement _element)
        {
            return _element is C.NumberingCache;
        }

        /// <summary>
        /// Check if the element is the point
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override bool CheckIsPointer(OpenXmlElement _element)
        {
            return _element is C.NumericPoint;
        }

        /// <summary>
        /// Create a new data point
        /// </summary>
        /// <returns></returns>
        protected override OpenXmlElement CreatePointer()
        {
            return new C.NumericPoint();
        }
    }

    /// <summary>
    /// The class for manipulating the values of the string data
    /// </summary>
    public class FelisChartStringReference : FelisChartDataReferance
    {
        internal FelisChartStringReference(C.StringReference _element)
            : base(_element)
        {
        }

        /// <summary>
        /// Check if the element is the cache
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override bool CheckIsCache(OpenXmlElement _element)
        {
            return _element is C.StringCache;
        }

        /// <summary>
        /// Check if the element is the point
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override bool CheckIsPointer(OpenXmlElement _element)
        {
            return _element is C.StringPoint;
        }

        /// <summary>
        /// Create a new data point
        /// </summary>
        /// <returns></returns>
        protected override OpenXmlElement CreatePointer()
        {
            return new C.StringPoint();
        }
    }

    /// <summary>
    /// The class for manipulating the valuse of the chart
    /// </summary>
    public class FelisChartValues : FelisChartNumberReference
    {
        /// <summary>
        /// The element of the values' container
        /// </summary>
        public readonly C.Values ValuesElement;

        internal FelisChartValues(C.Values _valsElement)
            : base(_valsElement.GetFirstChild<C.NumberReference>() ?? _valsElement.InsertElement(new C.NumberReference(), 0))
        {
            ValuesElement = _valsElement;
        }
    }

    /// <summary>
    /// The class fot manipulating the categories of the chart
    /// </summary>
    public class FelisChartCategories
    {
        /// <summary>
        /// The element of the categories' container
        /// </summary>
        public readonly C.CategoryAxisData CategoryElement;

        internal FelisChartCategories(C.CategoryAxisData _catElement)
        { 
            CategoryElement = _catElement;
            Reload();
        }

        private Lazy<FelisChartDataReferance>? reference;

        private FelisChartDataReferance ReferenceLoader()
        {
            var numRef = CategoryElement.GetFirstChild<C.NumberReference>();
            if (null != numRef)
            {
                return new FelisChartNumberReference(numRef);
            }

            var strRef = CategoryElement.GetFirstChild<C.StringReference>() ?? CategoryElement.InsertElement(new C.StringReference(), 0);
            return new FelisChartStringReference(strRef);
        }

        /// <summary>
        /// Reload the data of the categories
        /// </summary>
        public void Reload()
        {
            Interlocked.Exchange(ref reference, new Lazy<FelisChartDataReferance>(ReferenceLoader));
        }

        /// <summary>
        /// Get the values of the categories
        /// </summary>
        public object?[] Values => reference?.Value.Values ?? Array.Empty<string?>();

        /// <summary>
        /// Get or set the value of the category in the special index
        /// </summary>
        /// <param name="_index"></param>
        /// <returns></returns>
        public object? this[int _index]
        {
            get => reference?.Value[_index];
            set
            {
                if (null != reference)
                {
                    reference.Value[_index] = value;
                }
            }
        }

        /// <summary>
        /// Rewrite all the value of the categories
        /// </summary>
        /// <param name="_values"></param>
        public void ReWrite(IEnumerable<object?> _values)
        {
            var refObj = reference?.Value;

            if ((refObj is not FelisChartStringReference) && (null != refObj))
            {
                refObj.ReferenceElement.Remove();
                Reload();
                ReWrite(_values);
            }
            else
            {
                refObj?.ReWrite(_values);
            }
        }

        /// <summary>
        /// Rewrite all the value of the categories
        /// </summary>
        /// <param name="_values"></param>
        public void ReWrite(params object?[] _values)
        {
            ReWrite(_values as IEnumerable<object?>);
        }

        /// <summary>
        /// Get or set the reference of the data
        /// </summary>
        public (string? book, string? sheet, string start, string? end) DataReference
        {
            get
            {
                return reference?.Value.DataReference ?? (book: null, sheet: string.Empty, start: string.Empty, end: string.Empty);
            }
            set
            {
                if (null != reference)
                {
                    reference.Value.DataReference = value;
                }
            }
        }

        /// <summary>
        /// Remove the reference of the data
        /// </summary>
        public void RemoveReference()
        {
            reference?.Value.RemoveReference();
        }

        /// <summary>
        /// Set the reference to unknown
        /// </summary>
        /// <param name="_start">The loaction of the start data</param>
        /// <param name="_end">The location of the end data</param>
        public void SetUnknownReference(string _start, string _end)
        {
            reference?.Value.SetUnknownReference(_start, _end);
        }

        /// <summary>
        /// Get or set the format of the categories' data 
        /// </summary>
        public string? Format
        {
            get => reference?.Value?.Format;
            set
            {
                if (null != reference)
                {
                    reference.Value.Format = value;
                }
            }
        }
    }
}
