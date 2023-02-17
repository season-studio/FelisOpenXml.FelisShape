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
    /// Fill by pattern
    /// </summary>
    public class FelisPatternFillValue : FelisCompositeElement, IFelisFillValue
    {
        /// <summary>
        /// Create an empty pattern fill
        /// </summary>
        public FelisPatternFillValue() : this(new A.PatternFill()) { }

        internal FelisPatternFillValue(A.PatternFill _element) 
            : base(_element)
        {
            
        }

        /// <summary>
        /// The type of the fill element
        /// </summary>
        public Type? SDKElementType => typeof(A.PatternFill);
    }
}
