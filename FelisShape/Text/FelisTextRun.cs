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
    /// The class of the portion of the text
    /// </summary>
    public class FelisTextRun : FelisCompositeElement
    {
        internal FelisTextRun(A.Run _runElement)
            : base(_runElement)
        {
            
        }

        /// <summary>
        /// The text contained in the run
        /// </summary>
        public string? Text 
        {
            get
            {
                var textElement = Element.GetFirstChild<A.Text>();
                return textElement?.Text;
            }

            set
            {
                if (Element is A.Run runElement)
                {
                    if (null == runElement.Text)
                    {
                        runElement.Text = new A.Text();
                    }
                    var textElement = runElement.Text;
                    if (null != textElement)
                    {
                        textElement.Text = value ?? string.Empty;
                    }
                }
            }
        }

        /// <summary>
        /// Get the properties of the text
        /// </summary>
        public FelisTextProperties? Properties
        {
            get
            {
                var propElement = Element.GetFirstChild<A.RunProperties>() ?? new A.RunProperties();
                return new FelisTextProperties(propElement, Element);
            }
        }
    }

    /// <summary>
    /// Collection of the run
    /// </summary>
    public class FelisTextRunCollection : FelisModifiableCollection<A.Paragraph, A.Run, FelisTextRun>
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_container"></param>
        public FelisTextRunCollection(A.Paragraph _container) : base(_container)
        {
        }

        /// <summary>
        /// Boxing
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override FelisTextRun BoxingElement(A.Run _element)
        {
            return FSUtilities.SingletonAssignObject<FelisTextRun>(_element, (_) => new FelisTextRun(_element))!;
        }

        /// <summary>
        /// Create a new run
        /// </summary>
        /// <param name="_index"></param>
        /// <returns></returns>
        protected override A.Run CreateElement(int _index)
        {
            return new A.Run();
        }

        /// <summary>
        /// Unboxing
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        protected override OpenXmlElement UnboxingElement(FelisTextRun _element)
        {
            return _element.Element;
        }
    }
}
