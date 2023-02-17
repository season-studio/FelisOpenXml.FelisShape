using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The connection shape
    /// </summary>
    [FelisShapeClass(typeof(P.ConnectionShape), 
        NonVisualDrawingPropertiesChain = new[] { typeof(P.NonVisualConnectionShapeProperties) })
    ]
    public class FelisConnectionShape : FelisShape
    {
        internal FelisConnectionShape(P.ConnectionShape _element)
            : base(_element)
        { 
        }
    }
}
