using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using FelisOpenXml.FelisShape.Base;

namespace FelisOpenXml.FelisShape.Draw
{
    /// <summary>
    /// The fill class for the shape
    /// </summary>
    public class FelisFill : FelisUnderlingElement
    {
        internal FelisFill(OpenXmlCompositeElement _container, Action<object>? _submitter)
            : base(_container, _submitter)
        {
        }

        /// <summary>
        /// Reload the fill element
        /// </summary>
        protected override void Reload()
        {
            workElement = GetFillValueElement(Element);
        }

        /// <summary>
        /// The fill's value
        /// </summary>
        public IFelisFillValue? Value
        {
            get
            {
                var val = GetFillValueObject(workElement);
                if ((null == val) && (null == workElement) && (ContainerElement.Parent is P.Shape shape) && (shape.UseBackgroundFill?.Value ?? false))
                {
                    val = new FelisAsBackgroundFill();
                }
                return val;
            }

            set
            {
                if (null == value)
                {
                    if (workElement is not A.NoFill)
                    {
                        workElement?.Remove();
                        ContainerElement.AppendChild(new A.NoFill());
                    }
                }
                else
                {
                    if ((null != workElement) && (value.SDKElementType?.IsInstanceOfType(workElement) ?? false))
                    {
                        workElement.Remove();
                        workElement = null;
                    }

                    if (value is FelisAsBackgroundFill)
                    {
                        if (ContainerElement.Parent is P.Shape shape)
                        {
                            shape.UseBackgroundFill = true;
                        }
                    }
                    else if (null != value.Element)
                    {
                        ContainerElement.AppendChild(value.Element.Parent == null ? value.Element : value.Element.CloneNode(true));
                    }
                }
                Submit();
            }
        }

        /// <summary>
        /// Get the supported element of the fill for the given element
        /// </summary>
        /// <param name="_parentElement">The special element containing the fill</param>
        /// <returns></returns>
        public static OpenXmlElement? GetFillValueElement(OpenXmlElement? _parentElement)
        {
            if (null != _parentElement)
            {
                var list = _parentElement.ChildElements.Where(e =>
                {
                    return e switch
                    {
                        A.SolidFill => true,
                        A.GradientFill => true,
                        A.PatternFill => true,
                        A.NoFill => true,
                        A.BlipFill => true,
                        _ => false
                    };
                }).ToArray();
                if (list.Length > 1)
                {
                    foreach (var dropItem in list.Take(list.Length - 1))
                    {
                        dropItem.Remove();
                    }
                    return list[list.Length - 1];
                }
                else if (list.Length == 1)
                {
                    return list[0];
                }
            }

            return null;
        }

        /// <summary>
        /// Convert the fill element to the IFelisFillValue instance
        /// </summary>
        /// <param name="_element">The element of the fill</param>
        /// <returns></returns>
        public static IFelisFillValue? GetFillValueObject(OpenXmlElement? _element)
        {
            return _element switch
            {
                A.SolidFill solidFill => new FelisSolidFillValue(solidFill),
                A.GradientFill gradFill => new FelisGradientFillValue(gradFill),
                A.PatternFill patternFill => new FelisPatternFillValue(patternFill),
                A.BlipFill blipFill => new FelisBlipFillValue(blipFill),
                _ => null
            };
        }
    }
}
