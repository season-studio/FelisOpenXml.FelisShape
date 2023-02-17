using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Collections;
using System.Xml.Linq;
using System.Reflection;
using FelisOpenXml.FelisShape.Base;
using FelisOpenXml.FelisShape.Draw;
using FelisOpenXml.FelisShape.Text;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The information of the shape's rect
    /// </summary>
    public struct FelisShapeRect
    {
        /// <summary>
        /// x
        /// </summary>
        public long x;
        /// <summary>
        /// y
        /// </summary>
        public long y;
        /// <summary>
        /// cx
        /// </summary>
        public long cx;
        /// <summary>
        /// cy
        /// </summary>
        public long cy;
    }

    /// <summary>
    /// The class of the basic shape
    /// </summary>
    [FelisShapeClass(typeof(P.Shape), 
        NonVisualDrawingPropertiesChain = new[] { typeof(P.NonVisualShapeProperties) })
    ]
    public class FelisShape : FelisCompositeElement
    {
        internal FelisShape(OpenXmlCompositeElement _element)
            : base(_element)
        {

        }



        /// <summary>
        /// Get the non-visualable drawing properties element of the shape.
        /// A empty one will be created if there is no object contained in the shape.
        /// </summary>
        public P.NonVisualDrawingProperties NonVisualDrawingPropertiesSafe => (GetShapeClassAttribute(Element)?.GetNonVisualDrawingProperties(Element) ?? new P.NonVisualDrawingProperties());

        /// <summary>
        /// Get the non-visualable drawing properties element of the shape.
        /// </summary>
        public P.NonVisualDrawingProperties? NonVisualDrawingProperties => GetShapeClassAttribute(Element)?.GetNonVisualDrawingProperties(Element);

        /// <summary>
        /// Get the non-visualable properties element of the shape
        /// </summary>
        public OpenXmlElement? NonVisualProperties => GetShapeClassAttribute(Element)?.GetNonVisualProperties(Element);

        /// <summary>
        /// Get the properties element of the shape
        /// </summary>
        protected virtual OpenXmlCompositeElement? GetPropertiesElement(bool _forceOne)
        {
            var ret = Element.GetFirstChild<P.ShapeProperties>();
            if ((null == ret) && _forceOne)
            {
                ret = Element.AppendChild(new P.ShapeProperties());
            }
            return ret;
        }

        /// <summary>
        /// The ID of the shape
        /// </summary>
        public uint Id
        {
            get
            {
                return NonVisualDrawingPropertiesSafe.Id?.Value ?? 0;
            }

            set
            {
                var prop = NonVisualDrawingProperties;
                if (null != prop)
                {
                    prop.Id = value;
                }
            }
        }

        /// <summary>
        /// The name of the shape
        /// </summary>
        public string? Name
        {
            get
            {
                return NonVisualDrawingPropertiesSafe?.Name;
            }

            set
            {
                var prop = NonVisualDrawingProperties;
                if (null != prop)
                {
                    prop.Name = value;
                }
            }
        }

        /// <summary>
        /// Check if this shape is a place holder
        /// </summary>
        public bool IsPlaceHolder => NonVisualProperties?.Descendants<P.PlaceholderShape>().Any() ?? false;

        /// <summary>
        /// Get the type of the place holder
        /// </summary>
        public int? PlaceHolderType => (int?)(NonVisualProperties?.Descendants<P.PlaceholderShape>().FirstOrDefault()?.Type?.Value);

        /// <summary>
        /// Get the type of the place holder in text format
        /// </summary>
        public string? PlaceHolderTypeText => NonVisualProperties?.Descendants<P.PlaceholderShape>().FirstOrDefault()?.Type?.ToString();

        /// <summary>
        /// Get the shape of the place holder
        /// </summary>
        public FelisShape? PlaceHolderShape
        {
            get
            {
                var ph = NonVisualProperties?.Descendants<P.PlaceholderShape>().FirstOrDefault();
                if (null != ph)
                {
                    var phRefElements = Slide?.SDKPart?.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree?.ChildElements;
                    var phRef = phRefElements?.FirstOrDefault(e =>
                    {
                        if (e is OpenXmlCompositeElement ce)
                        {
                            var nvProps = GetShapeClassAttribute(ce)?.GetNonVisualProperties(ce);
                            var phRefProp = nvProps?.Descendants<P.PlaceholderShape>().FirstOrDefault();
                            if (null != phRefProp)
                            {
                                return Equals(ph.Type, phRefProp.Type) && Equals(ph.Index, phRefProp.Index);
                            }
                        }
                        return false;
                    });
                    if (phRef is OpenXmlCompositeElement phRefElement)
                    {
                        return CreateInstance(phRefElement);
                    }
                }
                return null;
            }
        }

        /// <summary>
        /// Get the a:xfrm or p:xfrm from a element
        /// </summary>
        /// <param name="_element"></param>
        /// <param name="_force"></param>
        /// <returns></returns>
        protected static OpenXmlCompositeElement? GetXFrm(OpenXmlCompositeElement _element, bool _force = false)
        {
            if (_element is P.GraphicFrame pGraphicFrame)
            {
                var xfrm = pGraphicFrame.Transform;
                if ((null == xfrm) && _force)
                {
                    xfrm = new P.Transform();
                    pGraphicFrame.Transform = xfrm;
                }
                return xfrm;
            }
            else if ((_element is P.GroupShape) || (_element is P.ShapeTree))
            {
                var prop = _element.GetFirstChild<P.GroupShapeProperties>();
                var xfrm = prop?.TransformGroup;
                if ((null == xfrm) && _force)
                {
                    xfrm = new A.TransformGroup();
                    if (null == prop)
                    {
                        prop = _element.InsertAt(new P.GroupShapeProperties(), 0);
                    }
                    prop.TransformGroup = xfrm;
                }
                return xfrm;
            }
            else
            {
                var prop = _element.GetFirstChild<P.ShapeProperties>();
                var xfrm = prop?.Transform2D;
                if ((null == xfrm) && _force)
                {
                    xfrm = new A.Transform2D();
                    if (null == prop)
                    {
                        prop = _element.InsertAt(new P.ShapeProperties(), 0);
                    }
                    prop.Transform2D = xfrm;
                }
                return xfrm;
            }
        }

        /// <summary>
        /// Get the a:xfrm or p:xfrm object
        /// </summary>
        /// <param name="_force">set true to create a empty one if there is not an existed one.</param>
        /// <returns></returns>
        protected OpenXmlCompositeElement? GetXFrm(bool _force = false)
        {
            return GetXFrm(Element, _force);
        }

        /// <summary>
        /// The rect parameter of the shape
        /// </summary>
        public FelisShapeRect Rect
        {
            get
            {
                var xfrm = GetXFrm() ?? PlaceHolderShape?.GetXFrm();
                long x = 0, y = 0, cx = 0, cy = 0;
                if (null != xfrm)
                {
                    var offset = xfrm.GetFirstChild<A.Offset>();
                    if (null != offset)
                    {
                        x = offset.X?.Value ?? 0;
                        y = offset.Y?.Value ?? 0;
                    }
                    var ext = xfrm.GetFirstChild<A.Extents>();
                    if (null != ext)
                    {
                        cx = ext.Cx?.Value ?? 0;
                        cy = ext.Cy?.Value ?? 0;
                    }
                }
                return new FelisShapeRect()
                {
                    x = x,
                    y = y,
                    cx = cx,
                    cy = cy
                };
            }

            set
            {
                var xfrm = GetXFrm(true);
                if (null != xfrm)
                {
                    var offset = xfrm.GetFirstChild<A.Offset>();
                    if (null == offset)
                    {
                        xfrm.AddChild(new A.Offset()
                        {
                            X = value.x,
                            Y = value.y
                        });
                    }
                    else
                    {
                        offset.X = value.x;
                        offset.Y = value.y;
                    }
                    var ext = xfrm.GetFirstChild<A.Extents>();
                    if (null == ext)
                    {
                        xfrm.AddChild(new A.Extents()
                        {
                            Cx = value.cx,
                            Cy = value.cy,
                        });
                    }
                    else
                    {
                        ext.Cx = value.cx;
                        ext.Cy = value.cy;
                    }
                }
            }
        }

        /// <summary>
        /// The rect parameter of the shape relative to the parent tree
        /// </summary>
        public FelisShapeRect RelativeRect
        {
            get
            {
                var absRect = Rect;
                if (Element.Parent is OpenXmlCompositeElement parent)
                {
                    var chOff = GetXFrm(parent)?.GetFirstChild<A.ChildOffset>();
                    if (null == chOff)
                    {
                        return absRect;
                    }
                    else
                    {
                        return new FelisShapeRect()
                        {
                            x = absRect.x - (chOff.X?.Value ?? 0),
                            y = absRect.y - (chOff.Y?.Value ?? 0),
                            cx = absRect.cx, 
                            cy = absRect.cy
                        };
                    }
                }
                else
                {
                    return absRect;
                }
            }

            set
            {
                if (Element.Parent is OpenXmlCompositeElement parent)
                {
                    var chOff = GetXFrm(parent)?.GetFirstChild<A.ChildOffset>();
                    if (null != chOff)
                    {
                        Rect = new FelisShapeRect()
                        {
                            x = value.x + (chOff.X?.Value ?? 0),
                            y = value.y + (chOff.Y?.Value ?? 0),
                            cx = value.cx,
                            cy = value.cy
                        };
                        return;
                    }
                }
                Rect = value;
            }
        }

        /// <summary>
        /// Get the parent shape containing this shape
        /// </summary>
        public FelisShape? Parent
        {
            get
            {
               return (Element.Parent is OpenXmlCompositeElement parentElement) ? FelisShape.CreateInstance(parentElement) : null;
            }
        }

        /// <summary>
        /// The slide which contains this shape
        /// </summary>
        public FelisSlide? Slide => FelisSlide.RetrospectToSlide(Element);

        /// <summary>
        /// Check if the shape contains text body
        /// </summary>
        public bool HasTextBody => Element.GetFirstChild<P.TextBody>() != null;

        /// <summary>
        /// Check if the shape contains text content
        /// </summary>
        public bool HasTextContent
        {
            get
            {
                return Element.GetFirstChild<P.TextBody>()?.Descendants<A.Text>().FirstOrDefault((text) => !string.IsNullOrEmpty(text.Text)) != null;
            }
        }

        /// <summary>
        /// Get the text body in the shape
        /// </summary>
        public FelisTextBody? TextBody
        {
            get
            {
                var textBodyElement = Element.GetFirstChild<P.TextBody>();
                return (null == textBodyElement) ? null : new FelisTextBody(textBodyElement);
            }
        }

        /// <summary>
        /// Get the fill of the shape
        /// </summary>
        public IFelisFillValue? Fill
        {
            get
            {
                var props = GetPropertiesElement(false);
                return (null == props) ? null : (new FelisFill(props, null)).Value;
            }

            set
            {
                var props = GetPropertiesElement(true);
                if (null != props)
                {
                    (new FelisFill(props, null)).Value = value;
                }
            }
        }

        #region 全局数据和行为
        internal static readonly Dictionary<Type, FelisShapeClassAttribute> ShapeClassMap = LoadShapeClassMap();

        /// <summary>
        /// Load the declared description attributes in this assembly
        /// </summary>
        /// <returns></returns>
        private static Dictionary<Type, FelisShapeClassAttribute> LoadShapeClassMap()
        {
            var map = new Dictionary<Type, FelisShapeClassAttribute>();

            var assembly = typeof(FelisShape).Assembly;
            foreach (var type in assembly.GetTypes())
            {
                if ((type == typeof(FelisShape)) || type.IsSubclassOf(typeof(FelisShape)))
                {
                    foreach (var attr in type.GetCustomAttributes(false))
                    {
                        if (attr is FelisShapeClassAttribute eAttr)
                        {
                            if (eAttr.ShapeType.IsSubclassOf(typeof(OpenXmlCompositeElement)))
                            {
                                FelisShapeClassAttribute? sAttr;
                                if (!map.TryGetValue(eAttr.ShapeType, out sAttr))
                                {
                                    sAttr = new FelisShapeClassAttribute(eAttr.ShapeType);
                                    map[eAttr.ShapeType] = sAttr;
                                }

                                if (null != sAttr)
                                {
                                    sAttr.Assign(eAttr);
                                    sAttr.TrySetCreator(type);
                                }
                            }
                        }
                    }
                }
            }

            return map;
        }

        /// <summary>
        /// Check is the input element is a shape element
        /// </summary>
        /// <param name="_shapeElement"></param>
        /// <returns></returns>
        public static bool IsShapeElement(object? _shapeElement)
        {
            return (null == _shapeElement) ? false : ShapeClassMap.ContainsKey(_shapeElement.GetType());
        }

        /// <summary>
        /// Create a instance of the FelisShape or the subclass from a given shape's element
        /// </summary>
        /// <param name="_shapeElement"></param>
        /// <returns></returns>
        public static FelisShape? CreateInstance(OpenXmlCompositeElement? _shapeElement)
        {
            return (null == _shapeElement) ? null : FSUtilities.SingletonAssignObject<FelisShape>(_shapeElement, (_) =>
            {
                if (ShapeClassMap.TryGetValue(_shapeElement.GetType(), out FelisShapeClassAttribute? attr))
                {
                    return attr.Creator?.Invoke(_shapeElement);
                }

                return null;
            });
        }

        /// <summary>
        /// Get the description attribute by the special shape's element
        /// </summary>
        /// <param name="_element"></param>
        /// <returns></returns>
        internal static FelisShapeClassAttribute? GetShapeClassAttribute(OpenXmlCompositeElement _element)
        {
            if (ShapeClassMap.TryGetValue(_element.GetType(), out FelisShapeClassAttribute? attr))
            {
                return attr;
            }
            return null;
        }

        /// <summary>
        /// Get the description attribute by the special class of the shape element
        /// </summary>
        /// <param name="_type"></param>
        /// <returns></returns>
        internal static FelisShapeClassAttribute? GetShapeClassAttribute(Type _type)
        {
            if (ShapeClassMap.TryGetValue(_type, out FelisShapeClassAttribute? attr))
            {
                return attr;
            }
            return null;
        }
        #endregion
    }
}
