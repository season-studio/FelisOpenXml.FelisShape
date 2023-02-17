using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FelisOpenXml.FelisShape.Base;

namespace FelisOpenXml.FelisShape.Text
{
    /// <summary>
    /// The class of the body of the text
    /// </summary>
    public class FelisTextBody : FelisCompositeElement
    {
        internal FelisTextBody(OpenXmlCompositeElement _textBodyElement)
            : base(_textBodyElement)
        {
            paragraphsCollection = new FelisTextParagraphCollection(_textBodyElement);
        }

        /// <summary>
        /// The cache of the paragraphs
        /// </summary>
        protected readonly FelisTextParagraphCollection paragraphsCollection;

        /// <summary>
        /// Get an iterator of all the paragraphs inside the text body
        /// </summary>
        public FelisTextParagraphCollection Paragraphs => paragraphsCollection;

        /// <summary>
        /// The text content inside the text body
        /// </summary>
        public string? Text
        {
            get
            {
                return string.Join(Environment.NewLine, Paragraphs.Select(e => e.Text));
            }

            set
            {
                var props = Paragraphs.FirstOrDefault()?.FirstTextProperties;
                Element.RemoveAllChildren<A.Paragraph>();
                if (null != value)
                {
                    Element.Append(value.Split('\n').Select(pText =>
                    {
                        var text = pText.TrimEnd('\r');
                        return new A.Paragraph(
                            (null != props) ? new A.Run(
                                props.Element.CloneNode(true),
                                new A.Text(text ?? string.Empty)
                            ) : new A.Run(
                                new A.Text(text ?? string.Empty)
                            )
                        );
                    }));
                }
            }
        }

        /// <summary>
        /// Get the first defined properties of text in the body.
        /// </summary>
        public FelisTextProperties? FirstTextProperties
        {
            get
            {
                foreach (var item in Element.Elements<A.Paragraph>())
                {
                    var prop = FelisTextParagraph.GetFirstTextProperties(item, false);
                    if (null != prop)
                    {
                        return prop;
                    }
                }
                var defRunProp = Element.GetFirstChild<A.ListStyle>()?.Elements<A.TextParagraphPropertiesType>().FirstOrDefault(e => e.Elements<A.DefaultRunProperties>().Any())?.Select(e => e.GetFirstChild<A.DefaultRunProperties>()).FirstOrDefault();
                if (null != defRunProp)
                {
                    return new FelisTextProperties(defRunProp, defRunProp.Parent);
                }
                return FelisShape.CreateInstance(Element.Parent as OpenXmlCompositeElement)?.PlaceHolderShape?.TextBody?.FirstTextProperties;
            }
        }
    }
}
