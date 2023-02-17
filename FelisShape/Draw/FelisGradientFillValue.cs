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
    /// Fill by gradient colors
    /// </summary>
    public class FelisGradientFillValue : FelisCompositeElement, IFelisFillValue
    {
        /// <summary>
        /// Create an empty gradient fill
        /// </summary>
        public FelisGradientFillValue() : this(new A.GradientFill()) { }

        internal FelisGradientFillValue(A.GradientFill _element)
            : base(_element) 
        { 
        }

        /// <summary>
        /// The type of the fill element
        /// </summary>
        public Type? SDKElementType => typeof(A.GradientFill);
    }
}
