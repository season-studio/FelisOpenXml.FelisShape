using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using FelisOpenXml.FelisShape.Draw;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace FelisOpenXml.FelisShape
{
    /// <summary>
    /// The picture shape class
    /// </summary>
    [FelisShapeClass(typeof(P.Picture), 
        NonVisualDrawingPropertiesChain = new[] { typeof(P.NonVisualPictureProperties) })
    ]
    public class FelisPicture : FelisShape
    {
        internal readonly FelisPictureBlip Blip;

        internal FelisPicture(P.Picture _element)
            : base(_element)
        {
            var blipElement = _element.GetFirstChild<P.BlipFill>();
            if (null == blipElement)
            {
                _element.AddChild(blipElement = new P.BlipFill(), false);
            }
            Blip = new FelisPictureBlip(blipElement);
        }

        /// <summary>
        /// Get the stretch object
        /// </summary>
        public FelisRelativeRect<A.FillRectangle> Stretch => Blip.Stretch;

        /// <summary>
        /// Get the displacement of the source
        /// </summary>
        public FelisRelativeRect<A.SourceRectangle> SourceDisplacement => Blip.SourceDisplacement;

        /// <summary>
        /// Get the data of the image
        /// </summary>
        /// <param name="_buffer">The buffer to receive the data. This argument can be a stream or an array of byte.</param>
        public void CopyTo(object _buffer)
        {
            Blip.CopyTo(_buffer);
        }

        /// <summary>
        /// Set a new image to this blip
        /// </summary>
        /// <param name="_source">The buffer containing the image. This argument can be a stream or an array of byte</param>
        /// <param name="_type">The type of the image. Such as "Bmp", "Png", and so on.</param>
        public void Set(object _source, string? _type)
        {
            Blip.Set(_source, _type);
        }
    }

    /// <summary>
    /// The class for p:blipFill element
    /// </summary>
    internal class FelisPictureBlip : FelisBlipBase<P.BlipFill>
    {
        internal FelisPictureBlip(P.BlipFill _element)
            : base(_element) 
        { 
        }
    }
}
