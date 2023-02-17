using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace FelisOpenXml.FelisShape.Draw
{
    /// <summary>
    /// Fill by the blip resource 
    /// </summary>
    public class FelisBlipFillValue : FelisBlipBase<A.BlipFill>, IFelisFillValue
    {
        /// <summary>
        /// Create a empty blip fill object
        /// </summary>
        public FelisBlipFillValue() : this(new A.BlipFill()) { }

        internal FelisBlipFillValue(A.BlipFill _element)
            : base(_element) 
        { 
        }

        /// <summary>
        /// The type of the blip element assigned to this object
        /// </summary>
        public Type? SDKElementType => typeof(A.BlipFill);
    }
}
