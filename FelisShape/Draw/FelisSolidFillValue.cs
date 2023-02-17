using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using FelisOpenXml.FelisShape.Base;
using A = DocumentFormat.OpenXml.Drawing;

namespace FelisOpenXml.FelisShape.Draw
{
    /// <summary>
    /// Fill by solid color
    /// </summary>
    public class FelisSolidFillValue : FelisCompositeElement, IFelisFillValue
    {
        /// <summary>
        /// Create an empty solid fill
        /// </summary>
        public FelisSolidFillValue() : this(new A.SolidFill()) { }

        internal FelisSolidFillValue(A.SolidFill _element)
            : base(_element)
        { 

        }

        /// <summary>
        /// The type of the fill element
        /// </summary>
        public Type? SDKElementType => typeof(A.SolidFill);

        /// <summary>
        /// The color of the fill
        /// </summary>
        public FelisColor? Color => FelisColor.Create(Element, null);
    }
}
