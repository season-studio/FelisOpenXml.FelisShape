using DocumentFormat.OpenXml;
using FelisOpenXml.FelisShape.Base;
using FelisOpenXml.FelisShape.Text;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The properties for paragraph
    /// </summary>
    public class FelisTextParagraphProperties : FelisCompositeElement
    {
        /// <summary>
        /// The element of the parent
        /// </summary>
        protected readonly OpenXmlElement? ParentElement;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_parentElement"></param>
        internal FelisTextParagraphProperties(A.TextParagraphPropertiesType _element, OpenXmlElement? _parentElement)
            : base(_element)
        {
            ParentElement = _parentElement;    
        }

        /// <summary>
        /// Submit the changing of the properties
        /// </summary>
        /// <returns></returns>
        public bool Submit()
        {
            try
            {
                if (null != ParentElement)
                {
                    if (!ParentElement.Contains(Element))
                    {
                        ParentElement.InsertElement(Element, 0);
                    }

                    return true;
                }
            }
            catch (Exception err)
            {
                Trace.TraceWarning(err.ToString());
            }

            return false;
        }

        /// <summary>
        /// Copy the properties from an other one
        /// </summary>
        /// <param name="_sourceProps"></param>
        /// <param name="_isPure"></param>
        public void CopyFrom(FelisTextParagraphProperties? _sourceProps, bool _isPure)
        {
            if (null != _sourceProps)
            {
                if (_isPure)
                {
                    Element.ClearAllAttributes();
                    Element.RemoveAllChildren();
                }

                foreach (var srcChild in _sourceProps.Element.Elements())
                {
                    var existedChildren = Element.Elements().Where(e => e.GetType() == srcChild.GetType()).ToArray();
                    if (existedChildren.Length > 0)
                    {
                        Element.InsertBefore(srcChild.CloneNode(true), existedChildren[0]);
                        foreach (var existedChild in existedChildren)
                        {
                            existedChild.Remove();
                        }
                    }
                    else
                    {
                        Element.Append(srcChild.CloneNode(true));
                    }
                }

                Element.SetAttributes(_sourceProps.Element.GetAttributes());
            }
        }

        /// <summary>
        /// The alignment
        /// </summary>
        public A.TextAlignmentTypeValues? Alignment
        {
            get
            {
                return (Element as A.TextParagraphPropertiesType)?.Alignment?.Value ?? null;
            }

            set
            {
                if (Element is A.TextParagraphPropertiesType props)
                {
                    props.Alignment = value;
                }
            }
        }
    }
}
