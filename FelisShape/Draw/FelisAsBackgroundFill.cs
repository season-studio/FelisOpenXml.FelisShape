using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FelisOpenXml.FelisShape.Draw
{
    /// <summary>
    /// Fill as the background of the slide
    /// </summary>
    public class FelisAsBackgroundFill: IFelisFillValue
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public FelisAsBackgroundFill() { }

        /// <summary>
        /// Invalid member of this object
        /// </summary>
        public Type? SDKElementType => null;

        /// <summary>
        /// Invalid member of this object
        /// </summary>
        public OpenXmlCompositeElement Element => null!;
    }
}
