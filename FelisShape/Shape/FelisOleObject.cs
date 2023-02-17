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
    /// The OLE object shape
    /// </summary>
    [FelisGraphicFrame(@"http://schemas.openxmlformats.org/presentationml/2006/ole")]
    public class FelisOleObject : FelisGraphicFrame
    {
        internal FelisOleObject(P.GraphicFrame _element)
            : base(_element)
        {

        }
    }
}
