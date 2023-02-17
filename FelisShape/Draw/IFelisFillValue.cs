using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FelisOpenXml.FelisShape.Draw
{
    /// <summary>
    /// Interface of the fill's value
    /// </summary>
    public interface IFelisFillValue
    {
        /// <summary>
        /// The fill element in Open XML SDK
        /// </summary>
        public OpenXmlCompositeElement Element { get; }

        /// <summary>
        /// The type of the fill element in Open XML SDK
        /// </summary>
        public Type? SDKElementType { get; }
    }
}
