using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The group shape
    /// </summary>
    [FelisShapeClass(typeof(P.GroupShape), 
        NonVisualDrawingPropertiesChain = new[] { typeof(P.NonVisualGroupShapeProperties) })
    ]
    public class FelisShapeGroup : FelisShape, IFelisShapeTree
    {
        internal FelisShapeGroup(P.GroupShape _shape)
            : base(_shape)
        { 
        }

        /// <summary>
        /// Get the properties element of the shape
        /// </summary>
        protected override OpenXmlCompositeElement? GetPropertiesElement(bool _forceOne)
        {
            var ret = Element.GetFirstChild<P.GroupShapeProperties>();
            if ((null == ret) && _forceOne)
            {
                Element.AddChild(ret = new P.GroupShapeProperties(), false);
            }
            return ret;
        }

        /// <summary>
        /// Check if there is any shape in the group
        /// </summary>
        public bool HasShapes => CheckHasShapes(Element);

        /// <summary>
        /// Get an iterator of the shapes in the group
        /// </summary>
        public IEnumerable<FelisShape> Shapes => GetChildrenShapes(Element);

        /// <summary>
        /// Check is there is any shape in the special element
        /// </summary>
        /// <param name="_element">The element which may contain some shapes</param>
        /// <returns>The result is true if there is any shape inside the given element.</returns>
        internal static bool CheckHasShapes(OpenXmlCompositeElement? _element)
        {
            if (null != _element)
            {
                foreach (var child in _element.ChildElements)
                {
                    if (FelisShape.IsShapeElement(child))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Get an iterator of the shapes inside the given element
        /// </summary>
        /// <param name="_element">The element which may contain some shapes</param>
        /// <returns>The iterator of the shapes inside the given element</returns>
        internal static IEnumerable<FelisShape> GetChildrenShapes(OpenXmlCompositeElement? _element)
        {
            if (null != _element)
            {
                foreach (var child in _element.ChildElements)
                {
                    var shape = (child is OpenXmlCompositeElement maybeShape) ? CreateInstance(maybeShape) : null;
                    if (null != shape)
                    {
                        yield return shape;
                    }
                }
            }
        }
    }
}
