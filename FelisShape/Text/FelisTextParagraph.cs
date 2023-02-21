using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using FelisOpenXml.FelisShape.Base;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace FelisOpenXml.FelisShape.Text
{
    /// <summary>
    /// The class of the paragraph
    /// </summary>
    public class FelisTextParagraph : FelisCompositeElement
    {
        internal FelisTextParagraph(A.Paragraph _paragraphElement)
            : base(_paragraphElement)
        {
            runsCollection = new FelisTextRunCollection(_paragraphElement);
        }

        /// <summary>
        /// The cache of the runs
        /// </summary>
        protected readonly FelisTextRunCollection runsCollection;

        /// <summary>
        /// The collection of the run is the paragraph
        /// </summary>
        public FelisTextRunCollection TextRuns => runsCollection;

        /// <summary>
        /// The text content in the paragraph
        /// </summary>
        public string? Text
        {
            get
            {
                return string.Concat(TextRuns.Select(e => e.Text));
            }

            set
            {
                Element.RemoveAllChildren<A.Run>();
                var props = FirstTextProperties;
                var run = (null != props) ? new A.Run(
                    props.Element.CloneNode(true),
                    new A.Text(value ?? string.Empty)
                ) : new A.Run(
                    new A.Text(value ?? string.Empty)
                );
                var epProps = Element.GetFirstChild<A.EndParagraphRunProperties>();
                if (null != epProps)
                {
                    Element.InsertBefore(run, epProps);
                }
                else
                {
                    Element.AddChild(run, false);
                }
            }
        }

        /// <summary>
        /// Get the first defined properties of text in the paragraph. An empty one will be created if there is not an existed one.
        /// </summary>
        public FelisTextProperties FirstTextProperties => GetFirstTextProperties(Element as A.Paragraph, true)!;

        /// <summary>
        /// Get the first defined properties of text in the paragraph
        /// </summary>
        /// <param name="_forceOne"></param>
        /// <returns></returns>
        public FelisTextProperties? GetFelisTextProperties(bool _forceOne)
        {
            return GetFirstTextProperties(Element as A.Paragraph, true);
        }

        /// <summary>
        /// The properties of the paragraph
        /// </summary>
        public FelisTextParagraphProperties? Properties
        {
            get
            {
                var props = Element.GetFirstChild<A.ParagraphProperties>();
                return (null != props) ? new FelisTextParagraphProperties(props, Element) : null;
            }
        }

        /// <summary>
        /// The properties of the paragraph.
        /// An empty one will be create if there is no existed on.
        /// </summary>
        public FelisTextParagraphProperties PropertiesSafe
        {
            get
            {
                var props = Element.GetFirstChild<A.ParagraphProperties>() ?? new A.ParagraphProperties();
                return new FelisTextParagraphProperties(props, Element);
            }
        }

        /// <summary>
        /// Get the properties of text defined as the level as the paragraph
        /// </summary>
        public FelisTextProperties ParagraphTextProperties
        {
            get
            {
                var endParaRPrElement = Element.GetFirstChild<A.EndParagraphRunProperties>();
                return new FelisTextProperties(endParaRPrElement ?? new A.EndParagraphRunProperties(), Element);
            }
        }

        /// <summary>
        /// Get the first defined properties of text in the given paragraph
        /// </summary>
        /// <param name="_paragraphElement">The element of the special paragraph</param>
        /// <param name="_forceOne">Set true for creating an empty one if there is not an existed properties in the given paragraph.</param>
        /// <returns></returns>
        internal static FelisTextProperties? GetFirstTextProperties(A.Paragraph? _paragraphElement, bool _forceOne)
        {
            if (null != _paragraphElement)
            {
                var runPropsElement = _paragraphElement.Descendants<A.RunProperties>().FirstOrDefault();
                if (null != runPropsElement)
                {
                    return new FelisTextProperties(runPropsElement, runPropsElement.Parent);
                }

                var endParaRPrElement = _paragraphElement.GetFirstChild<A.EndParagraphRunProperties>();
                return _forceOne ? new FelisTextProperties(endParaRPrElement ?? new A.EndParagraphRunProperties(), _paragraphElement)
                                 : ((null == endParaRPrElement) ? null : new FelisTextProperties(endParaRPrElement, _paragraphElement));
            }
            return null;
        }
    }

    /// <summary>
    /// Collection of the paragraphs
    /// </summary>
    public class FelisTextParagraphCollection : FelisModifiableCollection<OpenXmlCompositeElement, A.Paragraph, FelisTextParagraph>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_container"></param>
        public FelisTextParagraphCollection(OpenXmlCompositeElement _container) : base(_container)
        {
        }

        /// <summary>
        /// Boxing
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override FelisTextParagraph BoxingElement(A.Paragraph _element)
        {
            return FSUtilities.SingletonAssignObject<FelisTextParagraph>(_element, (_) => new FelisTextParagraph(_element))!;
        }

        /// <summary>
        /// Create paragraph
        /// </summary>
        /// <param name="_index"></param>
        /// <returns></returns>
        protected override A.Paragraph CreateElement(int _index)
        {
            return new A.Paragraph();
        }

        /// <summary>
        /// Unboxing
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        protected override OpenXmlElement UnboxingElement(FelisTextParagraph _element)
        {
            return _element.Element;
        }
    }
}
